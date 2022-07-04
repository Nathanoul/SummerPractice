from os import path, remove, mkdir, chdir
import win32com.client as win32
from MsgBoxListener import MsgBoxListener

class UseFile():

    def __init__(self,  filename:str, dirname:str = "tmp"):
        print("Opening excel file for macro executing...")

        # if path does not exist then making path with directory folder
        if not path.exists(dirname):
            mkdir(dirname)

        dirpath = dirname
        self._dirpath = dirpath
        self._fp = path.join(dirpath, filename)

        # checking if path to file exist
        if path.exists(self._fp):
            self.pathexists = True
        else:
            self.pathexists = False
            print("Cannot find file in the Input folder")


    def OpenFile(self):
        if self.pathexists:
            filepath = self._fp

            xl = win32.Dispatch("Excel.Application")  # creating COM object that grant access to the applications in VBA
            self._xl = xl

            wb = xl.Workbooks.Open(path.abspath(filepath))  # opening excel file
            self._wb = wb

            wb.Worksheets("массовый расчет").Activate()  # sheet with macros became active
            xl.DisplayAlerts = False
            wb.DoNotPromptForConvert = True
            wb.CheckCompatibility = False



    def CloseFile(self, editedfilename:str):
        if self.pathexists:
            dirpath = self._dirpath
            wb = self._wb
            xl = self._xl
            editedfilepath = path.join(dirpath, editedfilename)

            if path.exists(editedfilepath):
                remove(editedfilepath)  # Remove the older save
            wb.SaveAs(path.abspath(editedfilepath))

            wb.Close(True)
            xl.Quit()



    def ClickMassCalc(self, param1:int = 50,  param2:int = 1):
        if self.pathexists:
            xl = self._xl
            wb = self._wb

            wb.Worksheets("массовый расчет").Activate()  # sheet with macros became active

            sh = wb.Worksheets("массовый расчет")
            sh.Range("D3").Value = param1
            sh.Range("D4").Value = param2

            # creating lisener object to search and close popup
            # with name Microsoft Excel every 2 second
            listener = MsgBoxListener('Microsoft Excel', 2)
            listener.start()
            xl.Application.Run('click_mass_calc')  # execute "Запустить массовый расчет"
            listener.stop()
            message = listener.GetMessage()

            wb.Save()

            print("Mass calculating completed")
            print("Message:\n   " + message)  # print message



    def ClickSumOil(self):
        if self.pathexists:
            xl = self._xl
            wb = self._wb

            wb.Worksheets("массовый расчет").Activate()  # sheet with macros became active

            xl.Application.Run('click_sum_oilFields')  # execute "Суммировать добычу для месторождения"

            wb.Save()

            print("Oil fields completed")



    def ClickSumBushes(self):
        if self.pathexists:
            xl = self._xl
            wb = self._wb

            wb.Worksheets("массовый расчет").Activate()  # sheet with macros became active

            xl.Application.Run('click_sum_bushes')  # execute "Суммировать профиля добычи для каждого куста"

            wb.Save()

            print("Bushes completed")