#!/usr/bin/python

import sys
import xlrd

def usage_message(error):
    message = error + "\n"
    message += "You should use this script like this:\n"
    message += "./search-word.py work-you-want-to-search excel-file1 (excel-file2 ... excel-fileN)"

    return message

def parse_parameters(parameters):
    if len(parameters) < 3:
        print(usage_message("Not enought parameters"))
        sys.exit()

    return (parameters[1], parameters[2:])

def look_word_in_worksheet(word, workbook, worksheet_name):
    worksheet = workbook.sheet_by_name(worksheet_name)

    results = {}
    for current_row in range(worksheet.nrows):
        row = []

        for current_col in range(worksheet.ncols):
            cell_value = worksheet.cell_value(current_row, current_col)
            #FIXME Should we check cell_type?
            if cell_value == word:
                row.append(current_col)

        if row:
            results[current_row] = row

    return results

def look_word_in_file(word, filename):
    workbook = xlrd.open_workbook(filename)
    results = {}
    for worksheet_name in workbook.sheet_names():
        worksheet_result = look_word_in_worksheet(word, workbook, worksheet_name)
        if worksheet_result:
            results[worksheet_name] = worksheet_result

    return results

def show_results(results):
    message = ""

    for workbook_name, workbook_results in results.items():
        message += "Results found at " + workbook_name + ":\n"
        for worksheet_name, worksheet_results in workbook_results.items():
            message += "\tResults found at " + worksheet_name + " worksheet:\n"
            for row, col in worksheet_results.items():
                message += "\t\tResult found at row " + str(row + 1) + " (at cols " + str(col) + ")\n"

    return message


# Here starts the program
word, files = parse_parameters(sys.argv)

print "The word we are looking for: " + word
print "The files we are looking at: " + str(files)

results = {}
for workbook_name in files:
    results[workbook_name] = look_word_in_file(word, workbook_name)

print "\n" + show_results(results)
