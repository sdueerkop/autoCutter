#! /usr/bin/env python3

import os
import openpyxl
import logging

# Logging configuration
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

#TODO: Get root dir of the project
#TODO: Get list of folders
#TODO: For each folder chdir into the folder

#TODO: Create a 'parts' folder

#TODO: Extract the folder specific name (to be used in the output files of ffmpeg)


def readFromXlsx(filename):
    wb = openpyxl.load_workbook(filename, read_only=True)
    sheet = wb.get_sheet_by_name('Sheet1')
    for rowOfCellObjects in sheet['A2':'C13']:
        tempList = []
        for cellObj in rowOfCellObjects:
            tempList.append(cellObj.value)

        print(tuple(tempList))

#TODO: Cut the video with ffmpeg

def main():
    readFromXlsx('testWorkbook.xlsx')

if __name__ == "__main__":
    main()
