#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!
import argparse
import os
import sys
import re
import docx2txt

def check_args(args=None):
    parser = argparse.ArgumentParser(
        description="Search DOTM files in a directory for '$' character"
    )
    parser.add_argument("search_string", help="string to search for in DOTM files")
    parser.add_argument(
        "--dir",
        default=os.getcwd(),
        metavar="directory name",
        type=str,
        help="name of directory containing DOTM files",
    )
    arguments = parser.parse_args(args)
    return (arguments.search_string, arguments.dir)


def search_directory(dir_name, search_string):
    filenames = [
        os.path.join(dir_name, filename)
        for filename in os.listdir(dir_name)
        if filename.lower().endswith(".dotm")
    ]
    for filename in filenames:
        text = docx2txt.process(filename)
        match = re.search(r'[\s\S]{,40}\$[\s\S]{,40}', text)
        if match:
            print("Match found in file {}".format(filename))
            print("\t...{}...".format(match.group().encode('ascii', 'ignore')))


if __name__ == "__main__":
    search_string, directory = check_args(sys.argv[1:])
    search_directory(directory, search_string)
