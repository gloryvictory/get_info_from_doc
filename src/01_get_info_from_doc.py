#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
#   Author          :   Viacheslav Zamaraev
#   email           :   zamaraev@gmail.com
#   Script Name     : 01_get_info_from_doc.py
#   Created         : 13th Nov 2020
#   Last Modified	: 13th Тщм 2020
#   Version		    : 1.0
#   PIP             : pip install pywin32 and pip install pypiwin32
#   RESULT          :
# Modifications	: 1.1 -
#               : 1.2 -
#
# Description   : get some text

import os  # Load the Library Module
import os.path
import sys
import time
from sys import platform as _platform
from time import strftime  # Load just the strftime Module from Time
from datetime import datetime
import csv
import codecs
import logging
import win32com.client
#from pywin32 import win32com
#import pywin32
import uuid

import codecs
import os

import cfg


def doc2txt(folder_start=''):
    app = win32com.client.Dispatch('Word.Application')
    #app = pywin32.client.Dispatch('Word.Application')
    myDir = folder_start
    for subdir, dirs, files in os.walk(myDir):
        for file in files:
            file_path = subdir + os.path.sep + file
            if (file[-4:] != '.doc'):
                continue
            doc = app.Documents.Open(file_path)

            file = open('' + str(uuid.uuid4()) + '.txt', 'w+')
            ttt = str(doc.Content.Text)
            #file.write(ttt.encode('utf-8'))
            file.write(ttt)

            file.close()




# ---------------- do main --------------------------------
def main():
    time1 = datetime.now()
    print('Starting at :' + str(time1))

    dir_input = cfg.folder_in_win
    doc2txt(dir_input)

    #do_log_file()

    #do_multithreading(dir_input)


    time2 = datetime.now()
    print('Finishing at :' + str(time2))
    print('Total time : ' + str(time2 - time1))
    print('DONE !!!!')


if __name__ == '__main__':
    main()