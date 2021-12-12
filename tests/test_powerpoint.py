from pywinauto import Desktop,Application
import pywinauto
from app.powerpoint import Powerpoint

pptObj=Powerpoint()
class TestPowerpoint():
    def test_open_powerpoint(self):
        try:
            pptObj.launchppt()
        except Exception as e:
            print(e)
            raise

    def test_ChangeLayout(self):
        try:
            pptObj.changeSlideLayout()
        except Exception as e:
            print(e)
            raise

    def test_InsertPicture(self):
        try:
            picPath1="D:\Sophia\ATAGTR\SamplePic.PNG"
            pptObj.InsertPicture(picPath1)
        except Exception as e:
            print(e)
            raise

    def test_InsertSlideAndPicures(self):
        try:
            picPath2 = "D:\Sophia\ATAGTR\SamplePic1.PNG"
            picPath3 = "D:\Sophia\ATAGTR\SamplePic2.PNG"
            pptObj.InsertSlide()
            pptObj.InsertPicture(picPath2)
            pptObj.InsertSlide()
            pptObj.InsertPicture(picPath3)
        except Exception as e:
            print(e)
            raise

    def test_SlideShow(self):
        try:
            pptObj.StartSlideShow()
        except Exception as e:
            print(e)
            raise
