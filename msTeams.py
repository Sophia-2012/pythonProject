from pywinauto.application import Application

app=Application(backend='uia').start("C:\\Users\\sophi\\AppData\\Local\\Microsoft\\Teams\\current\\Teams.exe",timeout=20)
