import time

from pywinauto.application import Application

app=Application(backend='uia').start("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",timeout=20)

app.window(title_re='.*PowerPoint*.').wait('ready visible',timeout=20,retry_interval=1)
main = app.window(title_re='.*PowerPoint*.')
main.wrapper_object().maximize()
main.child_window(title="New", control_type="ListItem").click_input()
main.child_window(title="Blank Presentation", control_type="ListItem").click_input()

newmain=app.window(title_re='.*Presentation*.')
newmain.child_window(title="Layout", control_type="MenuItem").click_input()
newmain.child_window(title="Title and Content", control_type="ListItem").click_input()

time.sleep(5)
newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="Pictures", control_type="MenuItem").click_input()
newmain.child_window(title="This Device...", control_type="MenuItem").click_input()

newmain.child_window(title="Insert Picture",control_type="Window").wait('visible ready active')
open_dlg=newmain.child_window(title="Insert Picture",control_type="Window")
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").click_input()
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").type_keys("D:\Sophia\ATAGTR\SamplePic.PNG")
time.sleep(5)
open_dlg.OpenSplitButton.click_input()
time.sleep(5)

newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="New Slide", control_type="SplitButton").click_input()
time.sleep(5)

newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="Pictures", control_type="MenuItem").click_input()
newmain.child_window(title="This Device...", control_type="MenuItem").click_input()

newmain.child_window(title="Insert Picture",control_type="Window").wait('visible ready active')
open_dlg=newmain.child_window(title="Insert Picture",control_type="Window")
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").click_input()
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").type_keys("D:\Sophia\ATAGTR\SamplePic1.PNG")
time.sleep(5)
open_dlg.OpenSplitButton.click_input()
time.sleep(5)

newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="New Slide", control_type="SplitButton").click_input()
time.sleep(5)

newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="Pictures", control_type="MenuItem").click_input()
newmain.child_window(title="This Device...", control_type="MenuItem").click_input()

newmain.child_window(title="Insert Picture",control_type="Window").wait('visible ready active')
open_dlg=newmain.child_window(title="Insert Picture",control_type="Window")
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").click_input()
open_dlg.child_window(title="File name:", auto_id="1148", control_type="Edit").type_keys("D:\Sophia\ATAGTR\SamplePic2.PNG")
time.sleep(5)
open_dlg.OpenSplitButton.click_input()
time.sleep(5)

newmain.child_window(title="Insert", control_type="TabItem").click_input()
newmain.child_window(title="New Slide", control_type="SplitButton").click_input()
time.sleep(5)


newmain.child_window(title="Slide Show", control_type="TabItem", found_index=0).click_input()
newmain.child_window(title="From Beginning",control_type="Button", found_index=0).click_input()

#newmain.print_control_identifiers()