import pyparsing
import pandas as pd
import numpy as np
import json
import logging
import nltk, re, pprint
from nltk import word_tokenize
from decimal import Decimal

def file_finder():
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename
    
    root = Tk()
    root.withdraw()
    
    try:
        filepath = askopenfilename(title = "Select a text file to mine",filetypes = (("Text files","*.txt"),("all files","*.*")))

        if '.txt' in filepath.lower():
            filename = filepath.split('/')[-1].lower()
            filename = re.sub('\([0-9]*\)|.txt|- [0-9]*|-[0-9]*','',filename).strip()
            token = re.sub('\s','',filename)

            token_list = ['batchproof', 'chartofaccounts', 'ctdreg_othrswages','finalhierrollup', 'glmenu01', 'gmp11extcomp_mgcnt',
                          'gmp11extmgcnt_comp', 'hoursregister', 'mgtdtldrwnf_stck','mpcacctpayable', 'mpcacctrecvreport', 'mpcapclaimsreport',
                          'mpcgljournals', 'mpcmatandsupp', 'mpcpayroll', 'mpcprovliab','mpcrevenue', 'mpctreaworkreport', 'mpcworkcompreport','rrdreg_othrswages']

            if token in token_list:
                print('File format found\nConverting \'{}\' at \'location {}\' to an Excel file'.format(filename.capitalize(),filepath))
                return token, filepath, token_list

            else:
                print('File not found, please ensure that the file name is correct or the file is convertible. Check convertible file list to find out more')
                return None, None, None
        else:
            print('Please upload a text file for conversion')
            return None, None, None
    except:
        print('Something went wrong please re-run the code block')
        return None, None, None

def file_miner(token,filepath,token_list):
    try:
        if token in token_list:
            exec('{}(\'{}\')'.format(token,filepath))
    except Exception as e:
        print('Something went wrong please re-run the code block. \nError Code',e)

