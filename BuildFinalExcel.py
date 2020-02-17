from RowToObjects import RowMaker
from ExcelPathsGetter import FilesInDir
from DkfWinVersionOfficeVersionList import DkfWinVersionOfficeVersionList


class BuildFinalExcel:
    main_excel_pathname = ""
    dir_pdf_changed_to_excel = ""
    DkfWinOfficeMaker = ""
    dkf_win_office_list = []
    unique_elements_list = []

    def __init__(self, main_excel_pathname, dir_pdf_changed_to_excel):
        self.main_excel_pathname = main_excel_pathname
        self.dir_pdf_changed_to_excel = dir_pdf_changed_to_excel
        self.DkfWinOfficeMaker = DkfWinVersionOfficeVersionList()
        print("Builder finished")

    def build(self):
        print("Building final excel")
        row_objects, rows_with_bad_dkf = RowMaker.read_excel_file_to_sheet_and_making_object(self.main_excel_pathname)

        excel_paths = FilesInDir(self.dir_pdf_changed_to_excel).get_files_paths()
        for excel_path in excel_paths:
            for correct_dkf in row_objects:
                if excel_path.find(correct_dkf.get_id_dkf()) != -1:
                    print()
                    print("Correct path found   ", correct_dkf.get_id_dkf(), "    ", excel_path)
                    self.DkfWinOfficeMaker.add_found_pattern(correct_dkf.get_id_dkf(), excel_path)
        self.dkf_win_office_list = self.DkfWinOfficeMaker.get_actual_list()

    def show_dkfs_wins_offices_list(self):
        print()
        print("Final List:  ", self.dkf_win_office_list)

    def get_final_unique_elements_list(self):
        return self.dkf_win_office_list
