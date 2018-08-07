#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!
import glob
import sys
import re
import zipfile
import os

if __name__ == "__main__":
    search_text = sys.argv[1]
    if len(sys.argv) > 2:
        dotm_directory = sys.argv[2]
    else:
        dotm_directory = os.path.dirname(os.path.realpath(__file__))

dotm_files = glob.glob(dotm_directory + '/*.dotm')

print('Searching directory ' + dotm_directory + ' for text ' + search_text + ' ...')

matches = 0

for file in dotm_files:
    zip_ref = zipfile.ZipFile(file, 'r')
    file_content = zip_ref.open('word/document.xml').read()
    if search_text in file_content:
        match_index = file_content.index(search_text)
        print('Match found in file ' + file)
        print('   ...' + file_content[match_index-40:match_index+40])
        matches += 1

print(str(len(dotm_files)) + ' files searched\n' + str(matches) + ' total matches')