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

def look_word_in_file(word, filename):
    return [filename]

# Here starts the program
word, files = parse_parameters(sys.argv)

print "The word we are looking for: " + word
print "The files we are looking at: " + str(files)

results = {}
for workbook_name in files:
    results[workbook_name] = look_word_in_file(word, workbook_name)

print "The word has been found here:\n" + str(results)
