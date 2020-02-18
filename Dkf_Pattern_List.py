from SearchEngineExcel import SearchEngineExcel


class Dkf_Pattern_List:
    list_dkf_win_office = []

    def __init__(self, pattern_config_file):
        # Reading patterns list
        self.pattern_config_file = pattern_config_file

    def add_found_pattern(self, id_DKF, excel_Path):
        list_rows_with_patterns = self.get_patterns_from_search_engine(excel_Path)
        self.list_dkf_win_office.insert(len(self.list_dkf_win_office), [id_DKF, list_rows_with_patterns])

    def get_patterns_from_search_engine(self, excel_Path):
        SearchEngine = SearchEngineExcel(excel_Path, self.pattern_config_file)
        list_rows_with_patterns = SearchEngine.scan_file_for_patterns()
        return list_rows_with_patterns

    def get_actual_list(self):
        return self.list_dkf_win_office
