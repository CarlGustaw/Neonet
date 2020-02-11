import xlsxwriter


class ExcelWriter:
    workbook = None
    id_DKF = ""
    winVersion = ""
    officeVersion = ""
    listRowsWindows = []
    listRowsOffice = []
    pathToWriteExcel = 'C:/Users/dMichalczak/TestyPyKsiegowosc/KsiegowoscTest.xlsx'

    def createWorkBook(self):
        self.workbook = xlsxwriter.Workbook(self.pathToWriteExcel)
        return self.workbook

    def createWorkSheet(self):
        self.workbook = self.createWorkBook()
        worksheet = self.workbook.add_worksheet()
        return worksheet

    def makeExcel(self, listAsSheet):
        print("Number of elements in listAsSheet:   ", len(listAsSheet))
        # Create a workbook and add a worksheet.
        worksheet = self.createWorkSheet()

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Get information from row
        for element in listAsSheet:
            # Attribution values from list to appropriate variables
            self.id_DKF = element[0]
            self.winVersion = element[1]
            self.officeVersion = element[2]
            self.listRowsWindows = element[3]
            self.listRowsOffice = element[4]

            # Iterate over the data and write it out row by row.
            worksheet.write(row, col, self.id_DKF)
            worksheet.write(row, col + 1, self.winVersion)
            worksheet.write(row, col + 2, self.officeVersion)
            worksheet.write(row, col + 3, str(self.listRowsWindows))
            worksheet.write(row, col + 4, str(self.listRowsOffice))
            row += 1
        self.workbook.close()
