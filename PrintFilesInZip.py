import logging
from zipfile import ZipFile
import win32api
import win32print
import os
import time
import stat
import shutil
import win32timezone
from sys import argv
import subprocess

###################################################
#                   Logging
###################################################

# logging.debug('This is a debug message')
# logging.info('This is an info message')
# logging.warning('This is a warning message')
# logging.error('This is an error message')
# logging.critical('This is a critical message')

# logging.basicConfig(level=logging.DEBUG)

# logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
# logging.warning('This will get logged to a file')

# Specifying the name of the zip file and the normal file that the content will be moved to.
zipFolder = "FilesToBePrinted.zip"
normalFolder = "filesToBePrinted"


def unzip_and_extract_all(zipFolder, normalFolder):

    # Will extract all the files from a zip folder.
    # And place them in a normal folder.

    # open the zip file in read mode.
    with ZipFile(zipFolder, 'r') as zip:

        # extract all files
        zip.extractall(normalFolder)


def print_file(filename):
    # Will print filename with default printer
    win32api.ShellExecute(
        0,
        "print",
        filename,
        '"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )


def print_job_checker(array_of_all_docs):

    still_busy_printing = True

    while still_busy_printing:
        array_of_docs_in_printing_queue = []
        jobs = []

        # First get array of all the doc's in the printing queue
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
            flags, desc, name, comment = p
            phandle = win32print.OpenPrinter(name)
            print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
            if print_jobs:
                jobs.extend(list(print_jobs))
            for job in print_jobs:
                logging.debug("printer name => " + name)
                document = job["pDocument"]
                logging.debug("Document name => " + document)
                array_of_docs_in_printing_queue.append(document)
            win32print.ClosePrinter(phandle)

        logging.debug(array_of_docs_in_printing_queue)

        # Now we check to see if the are any of the doc's present
        # in both array_of_all_docs and array_of_docs_in_printing_queue.
        match_occurred = False
        for doc in array_of_all_docs:
            for printing_doc in array_of_docs_in_printing_queue:
                logging.debug(doc + " / " + printing_doc)

                if (doc == printing_doc):
                    logging.debug("still printing")
                    match_occurred = True

        if (match_occurred == False):
            still_busy_printing = False

        time.sleep(5)


def main(zipFolder, normalFolder):

    # Unzips folder and places documents in normalFolder.
    unzip_and_extract_all(zipFolder, normalFolder)

    # Array containing the name of all the documents.
    array_of_all_docs = []

    # Print each file in normal folder.
    for filename in os.listdir(normalFolder):
        logging.debug(filename)
        array_of_all_docs.append(filename)
        filePath = normalFolder + "\\" + filename
        logging.debug(filePath)
        print_file(filePath)

    logging.debug(array_of_all_docs)

    time.sleep(10)

    # Check How the printing is going.
    # Once all the doc's have been print we'll move on to the
    # clean-up stage.
    logging.debug("Running print_job_checker(array_of_all_docs)")
    print_job_checker(array_of_all_docs)

    logging.debug("can now start cleaning...")
    # time.sleep(5)

    # The Clean-up stage:
    # Remove zipFolder
    os.chmod(zipFolder, stat.S_IWRITE)
    os.remove(zipFolder)
    logging.debug("Removed zipFolder")

    # Remove normalFolder
    os.chmod(normalFolder, stat.S_IWRITE)
    shutil.rmtree(normalFolder)
    logging.debug("Removed normalFolder")

    # Now the program will delete itself.
    # os.remove(argv[0])

    # Create a bat file that will delete this executable.

    bat_file = "cleanup.bat"
    logging.debug("Creating cleanup.bat file")

    f = open(bat_file, "a")
    f.write("TASKKILL /IM PrintFilesInZip.exe /f\n")
    f.write("DEL " + "PrintFilesInZip.exe" + "\n")
    f.write('start /b "" cmd /c del "%~f0"&exit /b')
    f.close()

    # Run bat file
    logging.debug("Running cleanup.bat file")
    subprocess.call([bat_file])

main(zipFolder, normalFolder)
