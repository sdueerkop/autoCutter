#! /usr/bin/env python3

import os, glob
import openpyxl
import logging
import subprocess

#TODO: Handle for empty directories and empty xlsx-files

# Logging configuration
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

def readFromXlsx(filename):
    wb = openpyxl.load_workbook(filename, read_only=True)
    sheet = wb.get_sheet_by_name('Tabelle1')
    list_to_return = []
    for rowOfCellObjects in sheet['A2':'C18']:
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
        if start != 'x' and start != 'X' and end != 'x' and end != 'X' and start != None and end != None:
            convList = []
            #logging.debug("Aufgabe: {0}\nStartpunkt: {1}\nEnde: {2}".format(i,start,end))
            convList.append(str(i))
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
        else:
            logging.info("Kein Wert für Aufgabe {}".format(i))
        #logging.debug("Tuple of converted values: {}".format(tuple(convList)))
        list_to_return.append(tuple(convList))
    
    return tuple(sorted(set(list_to_return), key=lambda x: x[0]))

# Cut the video with ffmpeg

def callFfmpeg(list_of_tuples):
    # Input-Datei finden
    mp4Files = [f for f in glob.glob("*.MP4")]
    try:
        os.path.isfile(mp4Files[0])
        input_file = mp4Files[0]
        output_core = mp4Files[0].strip("Erhebung_I_")
        for t in list_of_tuples:
            # Output-Namen konstruieren
            output = "{}{}_{}".format('parts/',t[0],output_core)
            if os.path.isfile(output):
                logging.info("File {} exists. Skipping.".format(output))
                continue
            # Commando konstruieren
            cmd = 'ffmpeg -filter_complex highpass=f=400,lowpass=f=5800 -i {} -ss {} -to {} -async 1 {}'.format(input_file,t[1],t[2],output)
            cmd_as_list = cmd.split()
            subprocess.run(cmd_as_list)
    except:
        logging.exception("No media file in {}".format(os.getcwd()))

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
        
        # Check if media file exists
        if not os.path.isfile("Erhebung_I_{}.MP4".format(d)):
            logging.debug("Filename should be: Erhebung_I_{}.MP4".format(d))
            logging.info("No media file named in {}!".format(os.getcwd()))
            continue

        # Create a 'parts' folder
        try:
            os.mkdir('parts')
        except FileExistsError:
            logging.debug("Folder 'parts' exists in {}".format(os.getcwd())) 

        #Find xlsx files
        xlsx = [f for f in glob.glob("*.xlsx") and glob.glob("Timestamp_*")]
        logging.debug("In dir {} is {}".format(os.getcwd(), xlsx[0]))

        #Get the numbers from the excel sheet
        for x in xlsx:
            logging.debug("Filename is {}".format(x))
            list_of_times = readFromXlsx(x)
            converted = convertMinutes(list_of_times)
            callFfmpeg(converted)

if __name__ == "__main__":
    main()
