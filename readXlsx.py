#! /usr/bin/env python3

import os, glob
import openpyxl
import logging

# Logging configuration
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

def readFromXlsx(filename):
    wb = openpyxl.load_workbook(filename, read_only=True)
    sheet = wb.get_sheet_by_name('Sheet1')
    for rowOfCellObjects in sheet['A2':'C13']:
        tempList = []
        for cellObj in rowOfCellObjects:
            tempList.append(cellObj.value)
        logging.debug(tuple(tempList))

#TODO: Cut the video with ffmpeg

def main():
    # Get root dir of the project
    src_dir = input("Please enter the path to your video directories: ")
    logging.debug("Source dir set to: {}".format(src_dir))

    # Get list of folders
    list_of_dirs = next(os.walk(src_dir))[1]
    logging.debug(list_of_dirs)

    # For each folder chdir into the folder
    for d in list_of_dirs:
        os.chdir(os.path.join(src_dir,d))
        logging.debug("Current path is: {}".format(os.getcwd()))

        # Create a 'parts' folder
        try:
            os.mkdir('parts')
        except FileExistsError:
            logging.debug("Folder 'parts' exists in {}".format(os.getcwd())) 

        #Find xlsx files
        xlsx = [f for f in glob.glob("*.xlsx") and glob.glob("Timestamp_*")]
        logging.debug("In dir {} is {}".format(os.getcwd(), xlsx[0]))

        #: Extract the folder specific name (to be used in the output files of ffmpeg)
        output_name = xlsx[0].strip(".xlsx")
        output_name = output_name.strip("Timestamp_")
        logging.debug("Output name is: {}".format(output_name))

if __name__ == "__main__":
    main()
