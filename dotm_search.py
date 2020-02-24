#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "DavidRinSE"
from zipfile import ZipFile
import argparse
import sys
import os
import re


def search_dir(directory, searchTerm):
    filesSearched = 0
    filesFound = 0
    for root, _, files in os.walk(directory, topdown=False):
        for name in files:
            if name.endswith("dotm"):
                dotm = ZipFile(os.path.join(root, name))
                content = dotm.read('word/document.xml').decode('utf-8')
                filesSearched += 1
                if searchTerm in content:
                    filesFound += 1
                    index = content.index(searchTerm)
                    cleaned = re.sub('<(.|\n)*?>', '', content)
                    cleanedIndex = cleaned.index(searchTerm)
                    print("Match found while in file " + os.path.join(root, name))
                    print("   ...{}...\n   [cleaned]...{}...".format(content[index-40:index+40], cleaned[cleanedIndex-40:cleanedIndex+40]))
    print("Total dotm files searched: {}\nTotal dotm files found: {}".format(str(filesSearched), str(filesFound)))

def create_parser():
    parser = argparse.ArgumentParser(description="Searches directory with dotm files for the search term provided.")
    parser.add_argument('--dir', dest="directory", help='creates a summary file')
    parser.add_argument('search', help='search term')
    return parser

def main(args):
    parser = create_parser()
    ns = parser.parse_args(args)

    if not ns:
        parser.print_usage()
        sys.exit(1)

    searchTerm = ns.search
    directory = ns.directory if ns.directory else "."
    
    search_dir(directory, searchTerm)

if __name__ == '__main__':
    main(sys.argv[1:])