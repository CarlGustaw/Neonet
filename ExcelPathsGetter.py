import glob


class FilesInDir:
    excelPaths = []

    def __init__(self, DIRWITHPDFCHANGEDTOEXCEL):
        for filename in glob.glob(DIRWITHPDFCHANGEDTOEXCEL):
            self.excelPaths.append(filename)
        print("Excel files founded in dir")

    def getFilesPaths(self):
        return self.excelPaths
