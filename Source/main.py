from UseFile import UseFile
from ParseFile import ParseFile
import shutil
from os import chdir, listdir, path, remove, rmdir

if __name__ == '__main__':
    filename = "РБОТ_18022019_для ННГ.xlsm"

    # moving directory one folder back
    chdir("../")

    # creating necessary paths
    InputPath = "Input"
    OutputPath = "Output"
    tmpPath = "tmp"
    InputFilePath = path.join(InputPath, filename)
    tmpFilePath = path.join(tmpPath, filename)
    csvPath = path.join(tmpPath, "csvfiles")

    # copying files from Input folder to temporary folder
    shutil.copy2(InputFilePath, tmpFilePath)

    # creating execution object that allows to execute macro in excel file
    Executor = UseFile("РБОТ_18022019_для ННГ.xlsm")  # using РБОТ_18022019_для ННГ.xlsm

    Executor.OpenFile()  # opening file
    Executor.ClickMassCalc()  # execute "Запустить массовый расчет"
    Executor.ClickSumBushes()  # execute "Суммировать профиля добычи для каждого куста"
    Executor.ClickSumOil()  # execute "Суммировать добычу для месторождения"
    Executor.CloseFile("edited.xlsm")  # closing file and saving as edited.xlsm

    # creating parser object that allows to parse sheets in excel file
    Parser = ParseFile("edited.xlsm")  # using edited.xlsm

    # parse sheets 5, 6 and 7
    Parser.ParseSheet5("Sheet5.csv")
    Parser.ParseSheet6("Sheet6.csv")
    Parser.ParseSheet7("Sheet7.csv")

    # copying files to Output folder
    for f in listdir(csvPath):
        shutil.copy2(path.join(csvPath, f), path.join(OutputPath, f))

    # clearing temporary folder
    for f in listdir(csvPath):
        remove(path.join(csvPath, f))
    rmdir(csvPath)
    for f in listdir(tmpPath):
        remove(path.join(tmpPath, f))

    print("Done!")


