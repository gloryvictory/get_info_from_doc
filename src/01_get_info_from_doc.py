#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
#   Author          :   Viacheslav Zamaraev
#   email           :   zamaraev@gmail.com
#   Script Name     : 01_get_info_from_doc.py
#   Created         : 13th Nov 2020
#   Last Modified	: 13th Тщм 2020
#   Version		    : 1.0
#   PIP             : pip install
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

import cfg



# ---------------- do main --------------------------------
def main():
    time1 = datetime.now()
    print('Starting at :' + str(time1))

    dir_input = get_input_directory()

    #do_log_file()

    #do_multithreading(dir_input)


    time2 = datetime.now()
    print('Finishing at :' + str(time2))
    print('Total time : ' + str(time2 - time1))
    print('DONE !!!!')


if __name__ == '__main__':
    main()