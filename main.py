from data_structure import *
import argparse

# create parser
parser = argparse.ArgumentParser()
# add arguments to the parser
parser.add_argument("ltp_file")
# parse the arguments
args = parser.parse_args()
ltp___reportFile = args.ltp_file

report = ReportData()

Generator.file_parser_ltp(ltp___reportFile, report)

#optional (output from the datastructure)
Generator.list_test_cases(report)

Generator.format_excel_sheet(report, ltp___reportFile)