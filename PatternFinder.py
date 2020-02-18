from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter
import configparser
import ast

# Read config file
config = configparser.ConfigParser()
config.read('config.ini')
list_of_patterns = ast.literal_eval(config.get("PATTERNS", "pattern_config_file"))

# Write down values from config file
MAIN_EXCEL_PATH_NAME = config['INPUT']['MAIN_EXCEL_PATH_NAME']
DIR_WITH_PDF_CHANGED_TO_EXCEL = config['INPUT']['DIR_WITH_PDF_CHANGED_TO_EXCEL']
column_index_of_dkf = int(config['INPUT']['column_index_of_dkf'])
path_to_write_excel = config['OUTPUT']['path_to_write_excel']
pattern_config_file = list_of_patterns

# Search for patterns in excel files according to dkfs
builder = BuildFinalExcel(MAIN_EXCEL_PATH_NAME, DIR_WITH_PDF_CHANGED_TO_EXCEL, column_index_of_dkf, pattern_config_file)
builder.check_patterns_only_when_corresponding_pdf_to_excel_file_occur()

# Write down all founded patterns with dkf as new excel file
ToExcel = ExcelWriter(path_to_write_excel)
ToExcel.make_excel(builder.get_dkfs_patterns_list())

# Daniel Michalczak 18.02.2020
