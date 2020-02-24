#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Detrich"

import argparse
import os
import zipfile


parser = argparse.ArgumentParser()
parser.add_argument("-d", "--dir", help="Specify the directory")
parser.add_argument("-t", "--text", help="Specify text to search for")
args = parser.parse_args()


def decodeDotm(root, name):
    dotm = zipfile.ZipFile(os.path.join(root, name))
    content = dotm.read('word/document.xml').decode('utf-8')
    return content


def main(dir, text):
    for root, dirs, files in os.walk(dir, topdown=False):
        TotalFileCounter = 0
        TotalFoundCounter = 0
        print("Searching directory ./" + root + " for " + text + " ...")
        for name in files:
            if name.endswith(".dotm"):
                dotmFile = decodeDotm(root, name)
                TotalFileCounter += 1
                if text in dotmFile:
                    textIndex = dotmFile.index(text)
                    print("Found match in file ./" + os.path.join(root, name))
                    print("..." + dotmFile[textIndex-40:textIndex+40] + "...")
                    TotalFoundCounter += 1
        print("Total Matches found = " + str(TotalFoundCounter))
        print("Total Files Counted = " + str(TotalFileCounter))


if __name__ == '__main__':
    main(args.dir, args.text)
