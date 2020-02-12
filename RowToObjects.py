import ExcelReader
from RowRecord import RowRecord


class RowMaker:
    dataSheet = ""

    @staticmethod
    def readExcelFileToSheetAndMakingObject(mainExcelPathName, index_of_dkf_column, index_row_scan):
        dataSheet = ExcelReader.readExcelFile(mainExcelPathName)
        print("Main excel read successfully")
        listOfObjects = []
        rowsWithBadDKF = []
        how_many_has_plus = 0

        for rowNumber in range(index_row_scan, dataSheet.nrows - 1):
            # Check if row contain correct DKF
            if str(dataSheet.row_slice(rowNumber)[index_of_dkf_column].value).find("+") != -1 or \
                    dataSheet.row_slice(rowNumber)[index_of_dkf_column].value == "" or \
                    dataSheet.row_slice(rowNumber)[index_of_dkf_column].value == "brak skanu" or \
                    dataSheet.row_slice(rowNumber)[index_of_dkf_column].value == "FS-01WW/00032360/2014 - ZŚW 799/2014" or \
                    dataSheet.row_slice(rowNumber)[index_of_dkf_column].value == "ID:379199":
                if str(dataSheet.row_slice(rowNumber)[index_of_dkf_column].value).find("+") != -1:
                    how_many_has_plus += 1
                rowsWithBadDKF.append(rowNumber)
            else:
                listOfObjects.append(RowRecord(str(int(dataSheet.row_slice(rowNumber)[index_of_dkf_column].value))))

        print("Rows as objects add successfully", " Number of readed rows: ", dataSheet.nrows - 1)
        print("Number of rows with good DKF:  ", len(listOfObjects))
        print("Number of rows with bad DKF:  ", len(rowsWithBadDKF))
        print("Has plus in DKF:  ", how_many_has_plus)
        print()
        return listOfObjects, rowsWithBadDKF
