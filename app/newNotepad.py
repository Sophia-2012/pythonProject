from pywinauto.application import Application

app=Application(backend='uia').start('notepad.exe',timeout=10).connect(title='Untitled - Notepad',timeout=10)

textEditor=app.UntitledNotepad.child_window(title="Text Editor", auto_id="15", control_type="Edit").wrapper_object()
textEditor.type_keys("Hello welcome to ATAGTR 2021", with_spaces=True)

FormatMenu=app.UntitledNotepad.child_window(title="Format", control_type="MenuItem").wrapper_object()
FormatMenu.click_input()


FontItem=app.UntitledNotepad.child_window(title="Font...", auto_id="33", control_type="MenuItem").wrapper_object()
FontItem.click_input()
FontDialog=app.UntitledNotepad.Font
FontType=FontDialog.ComboBox
# # app.UntitledNotepad.print_control_identifiers()
# FontList=FontDialog.child_window(title="Font:", control_type="ComboBox").wrapper_object()
# FontList.select("Arial")
#app.UntitledNotepad.print_control_identifiers()
#FontDialog.child_window(title="Font:", control_type="ComboBox").select("Arial")
FontType.select(1)
FontStyle=FontDialog.ComboBox2
FontStyle.select(2)
FontSize=FontDialog.ComboBox3
FontSize.select(3)
FontDialog.OKButton.click_input()

