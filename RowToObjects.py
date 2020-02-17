import ExcelReader
from RowRecord import RowRecord


class RowMaker:

    def __init__(self, main_excel_pathname, column_index_of_dkf):
        self.main_excel_pathname = main_excel_pathname
        self.column_index_of_dkf = column_index_of_dkf
        self.data_sheet = ExcelReader.read_excel_file(self.main_excel_pathname)
        print("Main excel read successfully")
        self.list_of_objects = []

    def make_object(self):
        for row_number in range(1, self.data_sheet.nrows - 1):
            self.list_of_objects.append(
                RowRecord(str(self.data_sheet.row_slice(row_number)[self.column_index_of_dkf].value)))

        print("Rows as objects add successfully", " Number of read rows: ", self.data_sheet.nrows - 1)
        print("Number of rows with good DKF:  ", len(self.list_of_objects))
        print()
        return self.list_of_objects
