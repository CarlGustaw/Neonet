import xlsxwriter


class ExcelWriter:
    workbook = None
    id_DKF = ""
    win_Version = ""
    office_Version = ""
    list_rows_windows = []
    list_rows_office = []

    def __init__(self, path_to_write_excel):
        self.path_to_write_excel = path_to_write_excel

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

        # Start from the second cell. Rows are at 1 and columns are zero indexed.
        row = 1
        col = 0

        # Set column Names
        worksheet.write(0, col, "id_dkf")
        worksheet.write(0, col + 1, "line_with_search_pattern")

        # Get information from row
        for element in listAsSheet:
            # Attribution values from list to appropriate variables
            self.id_DKF = element[0]
            self.list_rows_windows = element[1]

            # Iterate over the data and write it out row by row.
            worksheet.write(row, col, self.id_DKF)
            worksheet.write(row, col + 1, str(self.list_rows_windows))
            row += 1
        self.workbook.close()
