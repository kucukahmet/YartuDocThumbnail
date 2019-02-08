import os, subprocess
from unocon import uno_start
from io import StringIO
from wand.image import Image
from pdf2image import convert_from_path

try:
  basestring
except NameError:
  basestring = str

DEFAULT_WIDTH, DEFAULT_HEIGHT = 128, 128

class Yimage(object):

    def create(self, file, out_path, width = DEFAULT_WIDTH, height = DEFAULT_HEIGHT):
        image = convert_from_path(file, first_page=1, last_page=1, output_folder = out_path)

class YartuDocThumb(Yimage):

    def __init__(self):
        self.uno = uno_start("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

    def pdf_to_img(self, file, out_path = "", width = DEFAULT_WIDTH, height = DEFAULT_HEIGHT):
        return super(YartuDocThumb, self).create(file, width=width, height=height)

    def doc_to_img(self, file, out_path = "", width = DEFAULT_WIDTH, height = DEFAULT_HEIGHT):
        pdf_path = self.uno.export_to_pdf(file, out_path = "/tmp/")
        return super(YartuDocThumb, self).create(pdf_path, out_path, width=width, height=height)

    def doc_to_pdf(self, file, out_path = ""):
        try:
            pdf = self.uno.export_to_pdf(file, out_path)
        finally:
            self.uno.close()

thumblier = YartuDocThumb()
# thumblier.doc_to_img(file = "abc.docx", out_path = "/home/ahmet/Desktop/")
thumblier.doc_to_pdf(file = "abc.docx", out_path = "/home/ahmet/Desktop/")
