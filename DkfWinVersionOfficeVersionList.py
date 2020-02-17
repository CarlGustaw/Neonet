from SearchEngineExcel import SearchEngineExcel


class DkfWinVersionOfficeVersionList:
    list_dkf_win_office = []

    def __init__(self):
        print("Creating ExcelReader")

    def add_found_pattern(self, id_DKF, excel_Path):
        win_version, office_version, list_rows_windows, list_rows_office = self.get_patterns_from_search_engine(excel_Path)

        # Removing None values from list -> clear view
        list_rows_windows = [i for i in list_rows_windows if i]

        self.list_dkf_win_office.insert(len(self.list_dkf_win_office),
                                        [id_DKF, win_version, office_version, list_rows_windows, list_rows_office])

    def get_patterns_from_search_engine(self, excel_Path):
        SearchEngine = SearchEngineExcel(excel_Path)
        winVersion, officeVersion, listRowsWindows, listRowsOffice = SearchEngine.scan_file_for_patterns()
        return winVersion, officeVersion, listRowsWindows, listRowsOffice

    def get_actual_list(self):
        return self.list_dkf_win_office
