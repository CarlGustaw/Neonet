from unittest import TestCase
import pytest

from ExcelReader import readExcelFile


class Test(TestCase):
    def test_read_excel_file_invalidPath(self):
        invalidPath = ""
        with pytest.raises(FileNotFoundError) as excinfo:
            readExcelFile(invalidPath)
        assert str(excinfo.value.args[1]) == 'No such file or directory'

    def test_read_excel_file_correctPath(self):
        correctPath = "D:/Poligon_Python/TestExcelFile.xlsx"
        with pytest.raises(FileNotFoundError) as excinfo:
            readExcelFile(correctPath)
        assert str(excinfo.value.args[1]) == 'No such file or directory'