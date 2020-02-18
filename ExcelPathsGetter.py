import glob


class FilesInDir:
    excel_paths = []

    def __init__(self, DIR_WITH_PDF_CHANGED_TO_EXCEL):
        for filename in glob.glob(DIR_WITH_PDF_CHANGED_TO_EXCEL):
            self.excel_paths.append(filename)
        number_of_read_files = len(self.excel_paths)
        print(number_of_read_files, "Excel files founded in dir")

    def get_files_paths(self):
        return self.excel_paths
