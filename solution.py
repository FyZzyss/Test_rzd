import json
import xlrd
import os
from argparse import ArgumentParser
from pathlib import Path


def main():
    final_dict = {}
    original_path, original_filename = os.path.dirname(args.input), os.path.splitext(os.path.basename(args.input))[0]
    output = args.output or args.input
    output_path, output_filename = os.path.dirname(output) or original_path, \
                                   Path(output).stem or original_filename  # get path + filename
    if output_path:
        output_path = os.path.join(output_path, '')
    output_filename = output_filename + '.json'
    if os.path.isfile(output_path + output_filename):  # simple check if file exists
        print('File with this name already exists')
    else:
        rb = xlrd.open_workbook(args.input)  # open xls
        sheet_names = rb.sheet_names()
        for name in sheet_names:  # get all sheet names and do loop
            sheet = rb.sheet_by_name(name)
            final_dict[name] = []
            for row_num in range(1, sheet.nrows):
                tmp_dict = {}
                for i in range(len(sheet.row_values(row_num))):
                    if 'DATE' in sheet.cell(0, i).value and sheet.cell(row_num, i).value != '':
                        tmp_dict[sheet.cell(0, i).value] = xlrd.xldate.xldate_as_datetime(sheet.cell(row_num, i).value,
                                                                                          sheet.book.datemode).strftime(
                            "%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z'  # format excel date to normal date object
                    elif sheet.cell(row_num, i).value == '':
                        tmp_dict[sheet.cell(0, i).value] = 'null'  # change empty columns to null
                    else:
                        if isinstance(sheet.cell(row_num, i).value, float):  # convert float to int
                            if sheet.cell(row_num, i).value == int(sheet.cell(row_num, i).value):
                                tmp_dict[sheet.cell(0, i).value] = int(sheet.cell(row_num, i).value)
                            else:
                                tmp_dict[sheet.cell(0, i).value] = sheet.cell(row_num, i).value
                        else:
                            if '{' in sheet.cell(row_num, i).value:  # if string is JSON convert to JSON
                                tmp_dict[sheet.cell(0, i).value] = json.loads(sheet.cell(row_num, i).value)
                            else:
                                tmp_dict[sheet.cell(0, i).value] = sheet.cell(row_num, i).value

                final_dict[name].append(tmp_dict)

        with open(output_path + output_filename, 'w', encoding='utf-8') as f:
            f.write(json.dumps(final_dict, ensure_ascii=False, indent=4))


if __name__ == '__main__':
    parser = ArgumentParser(description="Test Task for RZD")
    parser.add_argument(
        "-i",
        "--input_file",
        help="Path to excel file",
        action="store",
        dest="input",
        required=True
    )
    parser.add_argument(
        "-o",
        "--output_file",
        help="Path to output file",
        action="store",
        dest="output",
        required=False
    )
    args = parser.parse_args()
    main()
