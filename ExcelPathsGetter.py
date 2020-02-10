import glob


class FilesInDir:
    excelPaths = []

    def __init__(self, DIRWITHPDFCHANGEDTOEXCEL):
        for filename in glob.glob(DIRWITHPDFCHANGEDTOEXCEL):
            self.excelPaths.append(filename)
        print("Excel files founded in dir")
        self.showPath()

    def getFilesPaths(self):
        return self.excelPaths

    def showPath(self):
        for path in self.excelPaths:
            print(path)
