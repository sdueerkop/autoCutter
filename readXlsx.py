#! /usr/bin/env python3

import os, glob
import openpyxl
import logging

# Logging configuration
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

def readFromXlsx(filename):
    wb = openpyxl.load_workbook(filename, read_only=True)
    sheet = wb.get_sheet_by_name('Tabelle1')
    list_to_return = []
    for rowOfCellObjects in sheet['A2':'C13']:
        tempList = []
        for cellObj in rowOfCellObjects:
            tempList.append(cellObj.value)
        logging.debug("ReadFromXlsx yields: {}".format(tuple(tempList)))
        list_to_return.append(tuple(tempList))
    return list_to_return

# Change format of timestamps from 0,X to 00:00:X
def convertMinutes(inputList):
    list_to_return = []
    for i,start,end in inputList:
        #logging.debug("Aufgabe: {0}\nStartpunkt: {1}\nEnde: {2}".format(i,start,end))
        convList = []
        convList.append(i)
        start = "{0:.2f}".format(float(start))
        #logging.debug("Value for start is {}".format(start))
        padded_start = "00:{:05.2f}".format(float(start))
        #logging.debug("Value for padded_start is {}".format(padded_start))
        padded_start = padded_start.replace('.',':')
        convList.append(padded_start)
        end = "{0:.2f}".format(end)
        #logging.debug("Value for end is {}".format(end))
        padded_end = "00:{:05.2f}".format(float(end))
        #logging.debug("Value for padded_end is {}".format(padded_end))
        padded_end = padded_end.replace('.',':')
        convList.append(padded_end)
        #logging.debug("Tuple of converted values: {}".format(tuple(convList)))
        list_to_return.append(tuple(convList))
    
    return tuple(list_to_return)

#TODO: Cut the video with ffmpeg
# ffmpeg -i INPUT_FILE.MP4 -ss 00:00:03 -t 00:00:08 -async 1 OUTPUT_FILE.mp4

def main():
    # Get root dir of the project
    src_dir = input("Please enter the path to your video directories: ")
    logging.debug("Source dir set to: {}".format(src_dir))

    # Get list of folders
    list_of_dirs = next(os.walk(src_dir))[1]
    #logging.debug(list_of_dirs)

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
#        logging.debug("In dir {} is {}".format(os.getcwd(), xlsx[0]))

        #: Extract the folder specific name (to be used in the output files of ffmpeg)
        output_name = xlsx[0].strip(".xlsx")
        output_name = output_name.strip("Timestamp_")
        #logging.debug("Output name is: {}".format(output_name))

        #Get the numbers from the excel sheet
        for x in xlsx:
            logging.debug("Filename is {}".format(x))
            list_of_times = readFromXlsx(x)
            print(list_of_times)
            converted = convertMinutes(list_of_times)
            print(converted)

if __name__ == "__main__":
    main()
