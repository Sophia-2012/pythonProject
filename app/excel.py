import time
from pywinauto.application import Application
import pytest
import os
import random
import pywinauto

class ExcelCode():

    def launch(self):
        try:
            self.app=Application(backend='uia').start("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",timeout=40)
            self.app.window(title='Excel').wait('ready visible',timeout=40,retry_interval=1)
            self.main=self.app.window(title='Excel')
            self.main.maximize()

        except Exception as e:
            print(e)
            raise

    def openSavedExcel(self):
        try:
            self.main.child_window(title="Open", control_type="ListItem").click_input()
            time.sleep(5)
            self.main.child_window(title="Browse", control_type="Button").click_input()
            time.sleep(5)
            self.main.child_window(title="Open", control_type="Window").wait('visible ready active')
            self.open_dlg = self.main.child_window(title="Open", control_type="Window")
            self.open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").click_input()
            self.open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").type_keys("D:\Sophia\ATAGTR\Sample.xlsx")
            time.sleep(5)
            self.open_dlg.OpenSplitButton.click_input()
            time.sleep(5)

            self.newmain=self.app.window(title_re='.*Sample*.')
            self.newmain.set_focus()

            if self.newmain.EnableEditingButton.exists(timeout=40,retry_interval=1)==True:
                self.newmain.EnableEditingButton.click_input()

        except Exception as e:
            print(e)
            raise

    def applyPivotTable(self):
        try:
            self.newmain.child_window(title="Insert", control_type="TabItem").click_input()
            self.newmain.child_window(title="PivotChart", control_type="SplitButton").click_input()
            self.newmain.child_window(title="PivotChart", control_type="Button").click_input()
            pywinauto.keyboard.send_keys('{ENTER}')
            self.newmain.child_window(title="Country", control_type="CheckBox").click_input()
            self.newmain.child_window(title="Units Sold", control_type="CheckBox").click_input()
            self.newmain.child_window(title="Sale Price", control_type="CheckBox").click_input()
            time.sleep(5)
        except Exception as e:
            print(e)
            raise

    def saveExcel(self):
        try:
            self.n = random.randint(0, 22)
            self.savePath="D:\Sophia\ATAGTR\TestOutput\Sample"+str(self.n)+".xlsx"
            self.newmain.FileButton.click_input()
            # self.newmain.child_window(title="File", control_type="TabItem").click_input()
            self.newmain.child_window(title="Save As", control_type="ListItem").click_input()
            self.newmain.child_window(title="Browse", control_type="Button").click_input()
            time.sleep(5)
            self.newmain.child_window(title="Save As", control_type="Window").wait('visible ready active')
            self.save_dlg = self.newmain.child_window(title="Save As", control_type="Window")
            self.save_dlg.child_window(title="File name:", auto_id="1001", control_type="Edit").click_input()
            #self.save_dlg.child_window(title="File name:", auto_id="1001", control_type="Edit").type_keys("")
            pywinauto.keyboard.send_keys('^a{BACKSPACE}')
            self.save_dlg.child_window(title="File name:", auto_id="1001", control_type="Edit").type_keys(self.savePath)
            time.sleep(5)
            self.save_dlg.SaveButton.click_input()
            time.sleep(5)

        except Exception as e:
            print(e)
            raise



