from tkinter import filedialog
import os, sys
import pandas as pd
import tkinter as tk
from xml.dom import minidom
from xml.etree import ElementTree


root = tk.Tk()

class ExcelImporter(object):

    def __init__(self):
        self.my_filetypes = [('Excel File', 'xlsx'), ('Excel file', 'xlx')]

    def import_excelfile(self):
        self.filename = filedialog.askopenfile(parent=root, initialdir=os.getcwd(),
                                               title="Please select a file:")
        xlsx = pd.ExcelFile(self.filename.name)
        print(self.filename.name)
        #df = pd.read_excel(xlsx,sheet_name='end_point_group')
        #for index, row in df.iterrows():
        #    if row['description'] == 'POB CONTAINER MANAGEMENT - TEST':
        #        print(row['description'])
        return xlsx

    def prettify(self,elem):
        """Return a pretty-printed XML string for the Element.
        """
        rough_string = ElementTree.tostring(elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")


