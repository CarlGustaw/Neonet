import xlsxwriter


class ExcelWriter:
    workbook = None
    id_DKF = ""
    win_Version = ""
    office_Version = ""
    list_rows_windows = []
    list_rows_office = []
    path_to_write_excel = 'C:/Users/dMichalczak/TestyPyKsiegowosc/KsiegowoscTest.xlsx'

    def create_workbook(self):
        self.workbook = xlsxwriter.Workbook(self.path_to_write_excel)
        return self.workbook

    def create_worksheet(self):
        self.workbook = self.create_workbook()
        worksheet = self.workbook.add_worksheet()
        return worksheet

    def make_excel(self, listAsSheet):
        print("Number of elements in listAsSheet:   ", len(listAsSheet))
        # Create a workbook and add a worksheet.
        worksheet = self.create_worksheet()

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Get information from row
        for element in listAsSheet:
            # Attribution values from list to appropriate variables
            self.id_DKF = element[0]
            self.win_Version = element[1]
            self.office_Version = element[2]
            self.list_rows_windows = element[3]
            self.list_rows_office = element[4]

            # Iterate over the data and write it out row by row.
            worksheet.write(row, col, self.id_DKF)
            worksheet.write(row, col + 1, self.win_Version)
            worksheet.write(row, col + 2, self.office_Version)
            worksheet.write(row, col + 3, str(self.list_rows_windows))
            worksheet.write(row, col + 4, str(self.list_rows_office))
            row += 1
        self.workbook.close()
