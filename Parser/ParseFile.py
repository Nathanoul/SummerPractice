'''
Parses last three sheet in edited xlsm file to csv format
methods ParseSheet5, ParseSheet6 and ParseSheet7 parse sheets 5, 6 and 7 accordingly
'''
from os import path, remove, mkdir, chdir, getcwd
import openpyxl
import csv
import pandas as pd


class ParseFile:

    def __init__(self, filename: str, dirname: str = "tmp"):
        print("Opening excel file for parsing...")

        # if path does not exist then making path with directory folder
        if not path.exists(dirname):
            mkdir(dirname)

        # if path does not exist then making path with csvfiles
        if not path.exists(path.join(dirname, "csvfiles")):
            mkdir(path.join(dirname, "csvfiles"))

        self._csvpath = path.join(dirname, "csvfiles")
        filepath = path.join(dirname, filename)

        # checking if path to file exist
        if path.exists(filepath):
            self.pathexists = True
        else:
            self.pathexists = False
            print("Cannot find file in the UserFile folder")

        self._excel = openpyxl.load_workbook(filepath)  # opening excel workbook

    def ParseSheet5(self, parsedfilename: str):
        if self.pathexists:
            csvpath = self._csvpath
            parsedfilepath = path.abspath(path.join(csvpath, parsedfilename))

            sheet = self._excel['Текущая']

            col = csv.writer(open(parsedfilepath,
                                  'w',
                                  newline="",
                                  encoding='utf-8'))

            for r in sheet.rows:
                # row by row write
                if r[0].value:
                    col.writerow([cell.value for cell in r])
                else:
                    break

            print("Sheet 5 parsing complete")

    def ParseSheet6(self, parsedfilename: str):
        if self.pathexists:
            csvpath = self._csvpath
            parsedfilepath = path.abspath(path.join(csvpath, parsedfilename))

            sheet = self._excel['КустТекущая']

            col = csv.writer(open(parsedfilepath,
                                  'w',
                                  newline="",
                                  encoding='utf-8'))

            for r in sheet.rows:
                # row by row write
                if r[0].value:
                    col.writerow([cell.value for cell in r])
                else:
                    break

            print("Sheet 6 parsing complete")

    def ParseSheet7(self, parsedfilename: str):
        if self.pathexists:
            csvpath = self._csvpath
            parsedfilepath = path.abspath(path.join(csvpath, parsedfilename))

            sheet = self._excel['МестВынгапуровское']

            col = csv.writer(open(parsedfilepath,
                                  'w',
                                  newline="",
                                  encoding='utf-8'))

            stay = True
            for r in sheet.rows:
                # row by row write
                if r[0].value:
                    col.writerow([cell.value for cell in r])
                else:
                    if not stay:
                        stay = False
                        break

            print("Sheet 7 parsing complete")
