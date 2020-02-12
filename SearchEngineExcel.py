import ExcelReader


class SearchEngineExcel:
    pattern_windows_found = False
    excel_path_name = ""
    data_sheet = ""
    office_version = ""
    win_version = ""
    win_dict = {
        "WXP": "Windows XP",
        "WXPP": "Windows XP Pro",
        "WV": "Windows Vista",
        "WVP": "Windows Vista Pro",
        "W7": "Windows 7",
        "W7P": "Windows 7 Pro",
        "W8": "Windows 8",
        "W8P": "Windows 8",
        "W10": "Windows 10 Home",
        "W10P": "Windows 10 Pro",
        "Weird": "Weird"
    }
    office_dict = {
        "2007": "Office 2007",
        "2010": "Office 2010",
        "2013": "Office 2013",
        "2016": "Office 2016"
    }

    patterns_for_windows = ["windows", "win", "vis", "vb", "wlo", "w7", "w8", "w10", "wxp", "winlo", "w|o", "winlopro"]
    patterns_to_avoid = ["wr", "wy", "wys", "wa", "wie", "wn", "wh", "wa"]
    patterns_for_office = ["2007", "2010", "2013", "2016"]
    patterns_for_win_xp = ["windows xp", "wxp", 'winxp', 'win xp', "xp"]
    patterns_for_win_vista = ["windows vista", "wv", "winvista", "win vista", "vis", "vista", "winv", "vistabusiness"]
    patterns_for_win_7 = ["windows 7", "w7", "win7", "win 7", "w7p"]
    patterns_for_win_8 = ["windows 8", "w8", 'win8', 'win 8']
    patterns_for_win_10 = ["windows 10", "w10", "win10", "win 10", "win1", "win|", "winlo", "w|o", "winlopro"]

    patterns_for_index_error = ["i", "I", "j", "J", "|", "L", "f", "F", "£"]

    error_message = "Nie udalo sie odczytac wersji"

    def __init__(self, excelPathName):
        self.excel_path_name = excelPathName
        self.data_sheet = ExcelReader.readExcelFile(self.excel_path_name)

    def ScanFileForPatterns(self):
        list_rows_office_value = []
        list_row_windows_value = []
        for row_number in range(0, self.data_sheet.nrows - 1):
            for cell in self.data_sheet.row_slice(row_number):
                cell_string_value = str.lower(str(cell.value))

                # Searching for office version and writing down whole rows
                if (cell_string_value.find('off') != -1
                        and cell_string_value.find("officejet") == -1
                        and cell_string_value.find("officepower") == -1
                        or cell_string_value.find("otllce") != -1):
                    list_rows_office_value.append(self.data_sheet.row_slice(row_number))

                    self.search_for_pattern_in(cell_string_value, self.patterns_for_office, "", "")

                # Searching for windows version and writing down whole rows
                for pattern in self.patterns_for_windows:
                    if cell_string_value.find(pattern) != -1:
                        list_row_windows_value.append(self.data_sheet.row_slice(row_number))

                # Searching  for windows version
                list_row_windows_value.append(self.search_engine_for_windows(cell_string_value, row_number))

        self.if_no_version_found_set_error_message()
        self.show_information_found_about_windows_and_office_version()
        return self.win_version, self.office_version, list_row_windows_value, list_rows_office_value

    def search_engine_for_windows(self, cell_string_value, currentRowNumber):

        # Searching if in line is any trace of word windows or vista (some excels don't read pdfs correctly)
        if (cell_string_value.find('windows') != -1
                or cell_string_value.find('win') != -1
                or cell_string_value.find('vis') != -1
                or cell_string_value.find('vb') != -1
                or cell_string_value.find("wlO") != -1
                or cell_string_value.find("w|o") != -1
                or cell_string_value.find("winlo") != -1):

            if self.pattern_windows_found == False:
                self.search_for_pattern_in(cell_string_value, self.patterns_for_win_xp, "WXPP", "WXP")
            if self.pattern_windows_found == False:
                self.search_for_pattern_in(cell_string_value, self.patterns_for_win_vista, "WVP", "WV")
            if self.pattern_windows_found == False:
                self.search_for_pattern_in(cell_string_value, self.patterns_for_win_7, "W7P", "W7")
            if self.pattern_windows_found == False:
                self.search_for_pattern_in(cell_string_value, self.patterns_for_win_8, "W8P", "W8")
            if self.pattern_windows_found == False:
                self.search_for_pattern_in(cell_string_value, self.patterns_for_win_10, "W10P", "W10")

    # Method take as cell value as argument, specify pattern list to search and two dictionary links to version type.
    def search_for_pattern_in(self, cell_string_value, pattern_list, dict_link_if_pro, dict_link_if_not_pro):
        # If no dictionary link was provided, turn on module for search Office version
        if dict_link_if_pro == "" and dict_link_if_not_pro == "":
            for pattern in pattern_list:
                if cell_string_value.find(pattern) != -1:
                    self.office_version = self.office_dict.get(pattern)
                    break
                else:
                    self.office_version = cell_string_value
        else:
            # If dictionary link was provided, turn on module for search Windows version
            for pattern in pattern_list:
                if cell_string_value.find(pattern) != -1:

                    self.search_if_version_is_professional_or_not(cell_string_value,
                                                                  dict_link_if_pro,
                                                                  dict_link_if_not_pro)
                    self.pattern_windows_found = True
                    break

    def search_if_version_is_professional_or_not(self, cell_string_value, dict_link_if_pro, dictLinkIfNotPro):
        if (cell_string_value.find('pro') != -1
                or cell_string_value.find("w7p") != -1
                or cell_string_value.find("p") != -1
                and cell_string_value.find('prod') == -1
                and cell_string_value.find("xp") == -1):
            self.win_version = self.win_dict.get(dict_link_if_pro)
        else:
            self.win_version = self.win_dict.get(dictLinkIfNotPro)

    def if_no_version_found_set_error_message(self):
        if self.win_version == "":
            self.win_version = self.error_message
        if self.office_version == "":
            self.office_version = self.error_message

    def show_information_found_about_windows_and_office_version(self):
        print("From SearchEngine: Office version: ", self.office_version, "   Windows Version:  ", self.win_version)
