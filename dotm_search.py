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
import argparse

# parser = argparse.ArgumentParser(description='Finding out what argparse does')
# parser.add_argument('--dir', help='I need this')
# parser.add_argument('string', help='I need this')
# args = parser.parse_args()
# if args.string:
#     print("yayy?")
#     print(args.dir)
#     print(args.string)


def get_files(string_to_find, search_directory):
    all_files = []
    input_path = search_directory
    string_to_find = string_to_find
    for filename in os.listdir(input_path):
        all_files.append(filename)
    print_files_with_string(all_files, input_path, string_to_find)


def print_files_with_string(all_files, input_path, string_to_find):
    count_files_including_string = 0
    for my_file in all_files:
        file_location = os.path.join(input_path, my_file)
        if zipfile.is_zipfile(file_location):
            count_files_including_string += decode_file(
                file_location, string_to_find)
    print("")
    print("Search Directory Files: {} ".format(input_path))
    print(len(all_files))
    print ("Files found with string: '{}' " .format(string_to_find))
    print count_files_including_string


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


def create_parser():
    parser = argparse.ArgumentParser(
        description='Finding out what argparse does')
    parser.add_argument('-d', '--dir', help='I need this', default=".")
    parser.add_argument('string', help='I need this')
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()

    if not args:
        parser.print_usage()
        sys.exit()

    get_files(args.string, args.dir)


if __name__ == '__main__':
    main()
