'''
main script
Works in the following algorythm
Creates necessary paths if they are not exist
Copying files from UserFile folder to the temporary folder
opening excel to file and executing macro's
opening excel file to parse last three sheets of data to csv
copying csv file to FinalData folder
clearing tmp folder
'''
from Executor import *
from Parser import *
from FileManager import *


def main():
    # creating necessary paths if they are not exist
    outputfolderpath = "FinalData"
    inputfolderpath = "UserFile"
    tmpfolderpath = "tmp"
    csvfilesfolderpath = "tmp/csvfiles"
    CreateFolders([tmpfolderpath, outputfolderpath, csvfilesfolderpath])

    # copying files from UserFile folder to the temporary folder
    Copy2Folder(inputfolderpath, tmpfolderpath)

    # creating execution object that allows to execute macro in excel file
    Executor = UseFile("РБОТ_18022019_для ННГ.xlsm")  # using РБОТ_18022019_для ННГ.xlsm

    Executor.OpenFile()  # opening file
    p1 = 48
    p2 = 2
    Executor.ClickMassCalc(p1, p2)  # execute "Запустить массовый расчет" with parameters p1 and p2
    Executor.ClickSumBushes()  # execute "Суммировать профиля добычи для каждого куста"
    Executor.ClickSumOil()  # execute "Суммировать добычу для месторождения"
    Executor.CloseFile("edited.xlsm")  # closing file and saving as edited.xlsm

    # creating parser object that allows to parse sheets in excel file
    Parser = ParseFile("edited.xlsm")  # using edited.xlsm

    # parsing sheets 5, 6 and 7
    Parser.ParseSheet5("Sheet5.csv")
    Parser.ParseSheet6("Sheet6.csv")
    Parser.ParseSheet7("Sheet7.csv")

    # copying files to the FinalData folder
    Copy2Folder(csvfilesfolderpath, outputfolderpath)

    # clearing temporary folder
    ClearFolder(tmpfolderpath)

    print("Done!")


if __name__ == '__main__':
    main()
