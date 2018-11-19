#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!
import argparse
import os
import sys
from oletools.olevba3 import (
    VBA_Parser,
    TYPE_OLE,
    TYPE_OpenXML,
    TYPE_Word2003_XML,
    TYPE_MHTML,
)


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
        filename
        for filename in os.listdir(dir_name)
        if filename.lower().endswith(".dotm")
    ]
    for filename in filenames:
        # with open(filename, "rb") as f_obj:
        #     f_data = f_obj.read()
        #     vbaparser = VBA_Parser(filename, f_data)
        #     # try:
        #     #     print(vbaparser.reveal())
        #     # except Exception:
        #     #     pass
        #     # finally:
        #     #     vbaparser.close()
        document = Document(filename)
        print(dir(document))
        document.close()


if __name__ == "__main__":
    search_string, directory = check_args(sys.argv[1:])
    search_directory(directory, search_string)
