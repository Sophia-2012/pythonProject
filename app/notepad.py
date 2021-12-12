import time
from pywinauto.application import Application

class Notepad:

    def launchNotepad(self):
        try:
            self.app=Application().start('notepad.exe',timeout=10)
            self.main=self.app.window(title_re='.*Untitled*.')

        except Exception as e:
            print(e)
            raise

    def typeNotes(self):
        try:
            self.main.Edit.type_keys("Welcome to ATAGTR 2021!",with_spaces=True)
            time.sleep(5)
            #self.main.print_control_identifiers()

            self.font_menu=self.main.menu_select('Format->Font...')
            self.font_dlg=self.app.Font
            self.font_type=self.font_dlg.ComboBox
            self.font_type.select('Arial')

            self.font_style = self.font_dlg.ComboBox2

            self.font_style.select('Bold')

            self.font_size = self.font_dlg.ComboBox3

            self.font_size.type_keys('18')

            self.font_dlg.OK.click()

        except Exception as e:
            print(e)
            raise

    # def formatChange(self):
    #     try:
    #         #self.main.Format
    #         #self.main.child_window(title="Format", control_type="MenuItem").click_input()
    #     except Exception as e:
    #         print(e)
    #         raise