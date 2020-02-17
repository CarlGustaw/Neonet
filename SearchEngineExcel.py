import ExcelReader


class SearchEngineExcel:

    def __init__(self, excelPathName, pattern_config_file):
        self.excel_path_name = excelPathName
        self.data_sheet = ExcelReader.read_excel_file(self.excel_path_name)
        self.list_rows_with_patterns = []
        self.pattern_config_file = pattern_config_file
        self.error_message = "Pattern not found"

    def scan_file_for_patterns(self):
        for row_number in range(0, self.data_sheet.nrows):
            for pattern in self.pattern_config_file:
                for cell in self.data_sheet.row_slice(row_number):
                    cell_string_value = str.lower(str(cell.value))

                    # Add cell as element in to the list if pattern was found
                    if cell_string_value.find(pattern) != -1:
                        self.list_rows_with_patterns.append(cell_string_value)

        self.if_no_pattern_found_set_error_message()
        return self.list_rows_with_patterns

    def if_no_pattern_found_set_error_message(self):
        if len(self.list_rows_with_patterns) == 0:
            self.list_rows_with_patterns = self.error_message
