# Import pywinauto Application class

from pywinauto.application import Application

app=Application().start('notepad.exe',timeout=10)
main_dlg = app.UntitledNotepad
main_dlg.wait('visible')

main_dlg.Edit.type_keys("Hello pywinauto! It’s a sample text",

                        with_spaces=True,

                        with_newlines=True,

                        pause=0.5,

                        with_tabs=True)

font_menu = main_dlg.menu_select('Format->Font...')

font_dlg = app.Font

font_type = font_dlg.ComboBox

font_type.select('Comic Sans MS')

font_style = font_dlg.ComboBox2

font_style.select('Bold')

font_size = font_dlg.ComboBox3

font_size.type_keys('18')

font_dlg.OK.click()

