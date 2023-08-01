//
//  LocalizationsDataSource.swift
//  LocalizationEditor
//
//  Created by Igor Kulman on 30/05/2018.
//  Copyright © 2018 Igor Kulman. All rights reserved.
//

import Cocoa
import Foundation
import os
import xlsxwriter

//https://github.com/mehulparmar4ever/XlsxReaderWriterSwift

typealias LocalizationsDataSourceData = ([String], String?, [LocalizationGroup])

enum Filter: Int, CaseIterable {
    case all
    case missing
}

/**
 Data source for the NSTableView with localizations
 */
final class LocalizationsDataSource: NSObject {
    // MARK: - Properties

    private let localizationProvider = LocalizationProvider()
    private var localizationGroups: [LocalizationGroup] = []
    private var selectedLocalizationGroup: LocalizationGroup?
    private var languagesCount = 0
    private var masterLocalization: Localization?

    /**
     Dictionary indexed by localization key on the first level and by language on the second level for easier access
     */
    private var data: [String: [String: LocalizationString?]] = [:]

    /**
     Keys for the consumer. Depend on applied filter.
     */
    private var filteredKeys: [String] = []

    // MARK: - Actions
    
    func setupFormatHeader(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let format = workbook_add_format(workbook)
        format_set_font_name(format, NSString(string: "Verdana").fileSystemRepresentation)
        format_set_bold(format)
        format_set_font_size(format, 18)
        //format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(format, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bg_color(format, UInt32(LXW_COLOR_BLACK.rawValue))
        format_set_font_color(format, UInt32(LXW_COLOR_WHITE.rawValue))
        return format
    }
    
    func setupFormatText(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let myformatNormal = workbook_add_format(workbook)
        format_set_font_name(myformatNormal, NSString(string: "Verdana").fileSystemRepresentation) //Verdana
        format_set_align(myformatNormal, UInt8(LXW_ALIGN_VERTICAL_DISTRIBUTED.rawValue))
        format_set_text_wrap(myformatNormal);
        return myformatNormal
    }
    
    func export(excel: URL, onCompletion:@escaping () -> Void) {
        DispatchQueue.global(qos: .background).async { [unowned self] in
            
            let path = NSString(string: excel.path).fileSystemRepresentation
            //Generate workbook
            let workbook = workbook_new(path)
            //Add one Sheet in workbook //You can add multiple sheet
            for localizationGroup in self.localizationGroups {
                let worksheet = workbook_add_worksheet(workbook, localizationGroup.name)
                
                
                let formatHeader = self.setupFormatHeader(using: workbook)
                let formatText = self.setupFormatText(using: workbook)
                worksheet_set_column(worksheet, 0, 0, 50, nil)
                worksheet_set_row(worksheet, 0, 40, formatHeader)
                worksheet_write_string(worksheet, 0, 0, "词条(key)", formatHeader)
                worksheet_freeze_panes(worksheet, 0, 1)
                var locKeys = Array<String>()

                for (index, localization) in localizationGroup.localizations.enumerated() {
                    worksheet_set_column(worksheet, lxw_col_t(index+1), lxw_col_t(0), 50, nil)
                    //NSLocale* locale = [[NSLocale alloc] initWithLocaleIdentifier:obj];
                    //NSLocale *zhCNLocale = [NSLocale localeWithLocaleIdentifier:@"zh-CN"];
                    // NSString *displayName = [zhCNLocale displayNameForKey:NSLocaleLanguageCode value:languageCode];
                    let locale = Locale(identifier: localization.language)
                    let zhCNLocale = Locale(identifier: "zh-Hans")
                    let displayName = zhCNLocale.localizedString(forIdentifier: localization.language) ?? ""
                    //if let languageCode = locale.language.languageCode?.identifier {
                        
                    //}
                    let title = displayName.count > 0 ? "\(displayName)(\(localization.language))" : localization.language
                    worksheet_write_string(worksheet, 0, lxw_col_t(index+1), title, formatHeader)
                    for (row, localizationString) in localization.translations.enumerated() {
                        if(locKeys.contains(localizationString.key))
                        {
                            if let existRow = locKeys.firstIndex(of: localizationString.key) {
                                worksheet_write_string(worksheet, lxw_row_t(existRow+1), lxw_col_t(index+1), localizationString.value, formatText)
                            }
                        }
                        else
                        {
                            worksheet_set_row(worksheet, lxw_row_t(row+1), 35, formatText)
                            worksheet_write_string(worksheet, lxw_row_t(row+1), 0, localizationString.key, formatText)
                            worksheet_write_string(worksheet, lxw_row_t(row+1), lxw_col_t(index+1), localizationString.value, formatText)
                            locKeys.append(localizationString.key)
                        }
                    }
                }
            }
            
            workbook_close(workbook)
            DispatchQueue.main.async {
                onCompletion()
            }
        }
    }

    /**
     Loads data for directory at given path

     - Parameter folder: directory path to start the search
     - Parameter onCompletion: callback with data
     */
    func load(folder: URL, onCompletion: @escaping (LocalizationsDataSourceData) -> Void) {
        DispatchQueue.global(qos: .background).async {
            let localizationGroups = self.localizationProvider.getLocalizations(url: folder)
            guard localizationGroups.count > 0, let group = localizationGroups.first(where: { $0.name == "Localizable.strings" }) ?? localizationGroups.first else {
                os_log("No localization data found", type: OSLogType.error)
                DispatchQueue.main.async {
                    onCompletion(([], nil, []))
                }
                return
            }

            self.localizationGroups = localizationGroups
            let languages = self.select(group: group)

            DispatchQueue.main.async {
                onCompletion((languages, group.name, localizationGroups))
            }
        }
    }

    /**
     Selects given localization group, converting its data to a more usable form and returning an array of available languages

     - Parameter group: group to select
     - Returns: an array of available languages
     */
    private func select(group: LocalizationGroup) -> [String] {
        selectedLocalizationGroup = group

        let localizations = group.localizations.sorted(by: { lhs, rhs in
            if lhs.language.lowercased() == "base" {
                return true
            }

            if rhs.language.lowercased() == "base" {
                return false
            }

            return lhs.translations.count > rhs.translations.count
        })
        masterLocalization = localizations.first
        languagesCount = group.localizations.count

        data = [:]
        for key in masterLocalization!.translations.map({ $0.key }) {
            data[key] = [:]
            for localization in localizations {
                data[key]![localization.language] = localization.translations.first(where: { $0.key == key })
            }
        }

        // making sure filteredKeys are computed
        filter(by: Filter.all, searchString: nil)

        return localizations.map({ $0.language })
    }

    /**
     Selects given group and gets available languages

     - Parameter group: group name
     - Returns: array of languages
     */
    func selectGroupAndGetLanguages(for group: String) -> [String] {
        let group = localizationGroups.first(where: { $0.name == group })!
        let languages = select(group: group)
        return languages
    }

    /**
     Filters the data by given filter and search string. Empty search string means all data us included.

     Filtering is done by setting the filteredKeys property. A key is included if it matches the search string or any of its translations matches.
     */
    func filter(by filter: Filter, searchString: String?) {
        os_log("Filtering by %@", type: OSLogType.debug, "\(filter)")

        // first use filter, missing translation is a translation that is missing in any language for the given key
        let data = filter == .all ? self.data: self.data.filter({ dict in
            return dict.value.keys.count != self.languagesCount || !dict.value.values.allSatisfy({ $0?.value.isEmpty == false })
        })

        // no search string, just use teh filtered data
        guard let searchString = searchString, !searchString.isEmpty else {
            filteredKeys = data.keys.map({ $0 }).sorted(by: { $0<$1 })
            return
        }

        os_log("Searching for %@", type: OSLogType.debug, searchString)

        var keys: [String] = []
        for (key, value) in data {
            // include if key matches (no need to check further)
            if key.normalized.contains(searchString.normalized) {
                keys.append(key)
                continue
            }

            // include if any of the translations matches
            if value.compactMap({ $0.value }).map({ $0.value }).contains(where: { $0.normalized.contains(searchString.normalized) }) {
                keys.append(key)
            }
        }

        // sorting because the dictionary does not keep the sort
        filteredKeys = keys.sorted(by: { $0<$1 })
    }

    /**
     Gets key for speficied row

     - Parameter row: row number
     - Returns: key if valid
     */
    func getKey(row: Int) -> String? {
        return row < filteredKeys.count ? filteredKeys[row] : nil
    }

    /**
     Gets the message for specified row

     - Parameter row: row number
     - Returns: message if any
     */
    func getMessage(row: Int) -> String? {
        guard let key = getKey(row: row), let part = data[key], let firstKey = part.keys.map({ $0 }).first  else {
            return nil
        }

        return part[firstKey]??.message
    }

    /**
     Gets localization for specified language and row. The language should be always valid. The localization might be missing, returning it with empty value in that case

     - Parameter language: language to get the localization for
     - Parameter row: row number
     - Returns: localization string
     */
    func getLocalization(language: String, row: Int) -> LocalizationString {
        guard let key = getKey(row: row) else {
            // should not happen but you never know
            fatalError("No key for given row")
        }

        guard let section = data[key], let data = section[language], let localization = data else {
            return LocalizationString(key: key, value: "", message: "")
        }

        return localization
    }

    /**
     Updates given localization values in given language

     - Parameter language: language to update
     - Parameter key: localization string key
     - Parameter value: new value for the localization string
     */
    func updateLocalization(language: String, key: String, with value: String, message: String?) {
        guard let localization = selectedLocalizationGroup?.localizations.first(where: { $0.language == language }) else {
            return
        }
        localizationProvider.updateLocalization(localization: localization, key: key, with: value, message: message)
    }

    /**
     Deletes given key from all the localizations

     - Parameter key: key to delete
     */
    func deleteLocalization(key: String) {
        guard let selectedLocalizationGroup = selectedLocalizationGroup else {
            return
        }

        selectedLocalizationGroup.localizations.forEach({ localization in
            self.localizationProvider.deleteKeyFromLocalization(localization: localization, key: key)
        })
        data.removeValue(forKey: key)
    }

    /**
     Adds new localization key with a message to all the localizations

     - Parameter key: key to add
     - Parameter message: message (optional)
     */
    func addLocalizationKey(key: String, message: String?) {
        guard let selectedLocalizationGroup = selectedLocalizationGroup else {
            return
        }

        selectedLocalizationGroup.localizations.forEach({ localization in
            let newTranslation = localizationProvider.addKeyToLocalization(localization: localization, key: key, message: message)
            data[key] = [localization.language: newTranslation]
        })
    }

    /**
     Returns row number for given key

     - Parameter key: key to check

     - Returns: row number (if any)
     */
    func getRowForKey(key: String) -> Int? {
        return filteredKeys.firstIndex(of: key)
    }
}

// MARK: - Delegate

extension LocalizationsDataSource: NSTableViewDataSource {
    func numberOfRows(in _: NSTableView) -> Int {
        return filteredKeys.count
    }
}
