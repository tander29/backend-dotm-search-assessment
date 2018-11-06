#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Given a directory path, this searches all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!
import sys
import os
# from docx import Document
import doctest
import zipfile
# print('test')
# print(os.getcwd())
# print(os.listdir(os.getcwd()))

# with zipfile.ZipFile('/Users/travisanderson/Documents/dailyAssignments/backend/backend-dotm-search-assessment/test/A997.dotm', 'r') as file:
#     with file.open('word/document.xml', 'r') as s:
#         print(s.read())

# print(zipfile.is_zipfile(
#     '/Users/travisanderson/Documents/dailyAssignments/backend/backend-dotm-search-assessment/test/test.dotm'))


def get_files(string_to_find, search_directory):
    all_files = []
    input_path = search_directory
    string_to_find = string_to_find
    for filename in os.listdir(input_path):
        all_files.append(filename)
    print_files_with_string(all_files, input_path, string_to_find)


def print_files_with_string(all_files, input_path, string_to_find):
    count = 0
    for my_file in all_files:
        file_location = os.path.join(input_path, my_file)
        if zipfile.is_zipfile(file_location):
            count += decode_file(file_location, string_to_find)
    print("")
    print ("Files found:")
    print count


def decode_file(file_location, string_to_find):
    with zipfile.ZipFile(file_location, 'r') as file:
        with file.open('word/document.xml', 'r') as s:
            text = s.read()

            if string_to_find in text:
                dollar_location = text.find(string_to_find)
                print(" ")
                print(os.path.basename(file_location))
                print(file_location)
                print (text[dollar_location - 40: dollar_location] +
                       text[dollar_location: dollar_location+40])
                print(" ")
                return 1
    return 0


def main():
    arguments_amount = len(sys.argv)
    search_directory = os.getcwd()
    if arguments_amount < 2 or arguments_amount > 4:
        print "-----Need to input arguments 1 req, 2 optional, other conditions fail i.e python dotm_search.py myword filedirectory-----"
        return
    if arguments_amount == 2:
        string_to_find = str(sys.argv[1])
    if arguments_amount == 4:
        string_to_find = str(sys.argv[3])
        search_directory = os.path.join(search_directory, sys.argv[2][2:])
    get_files(string_to_find, search_directory)


if __name__ == '__main__':
    main()
