from UseFile import UseFile
from ParseFile import ParseFile
import shutil
from os import chdir, listdir, path, remove, rmdir, mkdir

if __name__ == '__main__':
    filename = "РБОТ_18022019_для ННГ.xlsm"

    # moving directory one folder back
    chdir("../")

    # creating necessary paths if they are not existing
    InputPath = "Input"
    if not path.exists(InputPath):
        mkdir(InputPath)

    OutputPath = "Output"
    if not path.exists(OutputPath):
        mkdir(OutputPath)

    tmpPath = "tmp"
    if not path.exists(tmpPath):
        mkdir(tmpPath)
        
    InputFilePath = path.join(InputPath, filename)
    tmpFilePath = path.join(tmpPath, filename)
    csvPath = path.join(tmpPath, "csvfiles")

    # copying files from Input folder to the temporary folder
    shutil.copy2(InputFilePath, tmpFilePath)

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

    # copying files to the Output folder
    for f in listdir(csvPath):
        shutil.copy2(path.join(csvPath, f), path.join(OutputPath, f))

    # clearing temporary folder
    for f in listdir(csvPath):
        remove(path.join(csvPath, f))
    rmdir(csvPath)
    for f in listdir(tmpPath):
        remove(path.join(tmpPath, f))

    print("Done!")


