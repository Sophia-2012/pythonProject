import time
from pywinauto.application import Application

class Powerpoint:
    pptPath="C:\\Program Files\Microsoft Office\\root\\Office16\\POWERPNT.EXE"

    def launchppt(self):
        try:
            self.app = Application(backend='uia').start("C:\\Program Files\Microsoft Office\\root\\Office16\\POWERPNT.EXE",timeout=20)
            self.app.window(title_re='.*PowerPoint*.').wait('ready visible',timeout=40,retry_interval=1)
            self.main=self.app.window(title_re='.*PowerPoint*.')
            self.main.wrapper_object().maximize()
            self.main.child_window(title="New", control_type="ListItem").click_input()
            self.main.child_window(title="Blank Presentation", control_type="ListItem").click_input()
            self.newmain = self.app.window(title_re='.*Presentation*.')
        except Exception as e:
            print(e)
            raise

    def changeSlideLayout(self):
        try:
            self.newmain.child_window(title="Layout", control_type="MenuItem").click_input()
            self.newmain.child_window(title="Title and Content", control_type="ListItem").click_input()
            time.sleep(5)
        except Exception as e:
            print(e)
            raise

    def InsertPicture(self,path):
        try:
            self.newmain.child_window(title="Insert", control_type="TabItem").click_input()
            self.newmain.child_window(title="Pictures", control_type="MenuItem").click_input()
            self.newmain.child_window(title="This Device...", control_type="MenuItem").click_input()

            self.newmain.child_window(title="Insert Picture", control_type="Window").wait('visible ready active')
            self.open_dlg = self.newmain.child_window(title="Insert Picture", control_type="Window")
            self.open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").click_input()
            self.open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").type_keys(path)
            time.sleep(5)
            self.open_dlg.OpenSplitButton.click_input()
            time.sleep(5)
        except Exception as e:
            print(e)
            raise

    def InsertSlide(self):
        try:
            self.newmain.child_window(title="Insert", control_type="TabItem").click_input()
            self.newmain.child_window(title="New Slide", control_type="SplitButton").click_input()
            time.sleep(5)
        except Exception as e:
            print(e)
            raise

    def StartSlideShow(self):
        try:
            self.newmain.child_window(title="Slide Show", control_type="TabItem", found_index=0).click_input()
            self.newmain.child_window(title="From Beginning", control_type="Button", found_index=0).click_input()
        except Exception as e:
            print(e)
            raise



