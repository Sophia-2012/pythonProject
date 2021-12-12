import pywinauto
from app.notepad import Notepad

notepadObj=Notepad()

class TestNotepad():


    #Open Notepad
    def test_launchNotepad(self):
        try:
            notepadObj.launchNotepad()
        except Exception as e:
            print("Launch of notepad was not successful: "+ str(e))
            raise

    def test_writeNotes(self):
        try:
            notepadObj.typeNotes()
        except Exception as e:
            print("Typing on the notepad was not successful: "+ str(e))
            raise

    # def test_format(self):
    #     try:
    #         notepadObj.formatChange()
    #     except Exception as e:
    #         print("Format on the notepad was not successful: "+ str(e))
    #         raise