from RowToObjects import RowMaker
from ExcelPathsGetter import FilesInDir
from Dkf_Pattern_List import Dkf_Pattern_List


class BuildFinalExcel:

    def __init__(self, main_excel_pathname, dir_pdf_changed_to_excel, column_index_of_dkf, pattern_config_file):
        self.main_excel_pathname = main_excel_pathname
        self.dir_pdf_changed_to_excel = dir_pdf_changed_to_excel
        self.column_index_of_dkf = column_index_of_dkf

        self.Dkf_Pattern_Maker = Dkf_Pattern_List(pattern_config_file)
        self.row_maker = RowMaker(self.main_excel_pathname, self.column_index_of_dkf)
        self.row_objects = self.__make_row_objects()
        print("Row transformed into objects")
        self.excel_paths = FilesInDir(self.dir_pdf_changed_to_excel).get_files_paths()
        print("Read pdfs changed into excel files")

    def __make_row_objects(self):
        return self.row_maker.make_object()

    def check_patterns_only_when_corresponding_pdf_to_excel_file_occur(self):
        for row_object in self.row_objects:
            for excel_path in self.excel_paths:

                # Search for excel with corresponding dkf name
                if excel_path.find(row_object.get_dkf()) != -1:
                    self.Dkf_Pattern_Maker.add_found_pattern(row_object.get_dkf(), excel_path)
                    print("Patterns found in ", excel_path)
                    print()

    def get_dkfs_patterns_list(self):
        return self.Dkf_Pattern_Maker.get_actual_list()
