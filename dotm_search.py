#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!
import argparse
import os


def check_args(args=None):
    parser = argparse.ArgumentParser(
        description='Search DOTM files in a directory for \'$\' character')
    parser.add_argument('search_string', default=os.getcwd(), help='string to search for in DOTM files')
    parser.add_argument('--dir', metavar='directory name',
                        type=str, help='name of directory containing DOTM files')
    arguments = parser.parse_args(args)
    return (arguments.search_string, arguments.dir)

def search_directory(search_string, dir_name = '.'):
    

if __name__ == '__main__':
    search_string, directory = check_args(sys.argv[1:])
