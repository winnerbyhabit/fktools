#! /usr/bin/env python3

	

import tkinter as tk
from tkinter import filedialog

import urllib.request

import os

#class to act like argparse Namespace
class Namespace:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)

def choose_data_files():
    # Ein Fenster erstellen
    fenster = tk.Tk()
   
    faelle = choose_a_file(fenster,"Wähle die Fälle.xls Datei aus",[("Excel Dateien","*.xlsx")])
    
    personen = choose_a_file(fenster,"Wähle die Personen.xls Datei aus",[("Excel Dateien","*.xlsx")])
    
    
    
    return faelle, personen 

def choose_a_file(fenster,message,file_types):
    currdir = os.getcwd()
    tempdir = ''
    while not tempdir:
        tempdir = filedialog.askopenfilename(parent=fenster, initialdir=currdir, title=message, filetypes=file_types)
    return tempdir

def download_xlsx_parser(path):
    url = "https://github.com/dilshod/xlsx2csv/raw/master/xlsx2csv.py"
    # Download the file from `url` and save it locally under `file_name`:
    urllib.request.urlretrieve(url, path)
    
def xls_to_csv(xlsx_parser,xlsx_file,save_path):
    from xlsx2csv import Xlsx2csv as parser
    parser(xlsx_file, outputencoding="utf-8").convert(save_path,sheetid=2)

def remove_unnecessary_lines(file,skiplines):
    with open(file, 'r') as fin:
        data = fin.read().splitlines(True)
    with open(file, 'w') as fout:
        fout.writelines(data[skiplines:])

def start_analyze(faelle_csv,personen_csv,fachschaftenliste_md):
    import analyze as analyze
    args = Namespace(faelle_csv = faelle_csv, personen_csv=personen_csv,fachschaftenliste_md=fachschaftenliste_md)
    analyze.analyze(args)

def download_fachschaften_md(path):
    url = "https://raw.githubusercontent.com/HSZemi/vs-bonn/master/md/Ordnungen/FKGO.md"
    # Download the file from `url` and save it locally under `file_name`:
    urllib.request.urlretrieve(url, path)
    
    with open(path, 'r') as fin:
        data = fin.read()
        data = data.split('# Anlage Fachschaftenliste')[1]
    with open(path, 'w') as fout:
        fout.write(data)
        
if __name__ == "__main__":
    faelle_csv = 'faelle.csv'
    personen_csv = 'personen.csv'
    fachschaften_md = 'fachschaften_md.md'
    
    
    faelle,personen = choose_data_files()
    download_xlsx_parser("xlsx2csv.py")
    download_fachschaften_md(fachschaften_md)
    #convert fälle
    xls_to_csv("xlsx2csv.py",faelle,faelle_csv)
    
    #convert personen
    xls_to_csv("xlsx2csv.py",personen,personen_csv)
    
    remove_unnecessary_lines("faelle.csv",5)
    remove_unnecessary_lines("personen.csv",5)
    
    start_analyze(faelle_csv,personen_csv,fachschaften_md)
