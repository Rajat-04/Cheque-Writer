# Cheque-Writer
A desktop application which lets user print multiple cheques without having to manually enter data, just fetch from excel sheet.

Python Modules Required:

from num2words import num2words 
from PyPDF2 import PdfMerger

from tkinter import * 
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter import Button
from tkinter import Checkbutton
from tkinter import Label
from tkinter import Entry
from tkinter import messagebox

import datetime
from datetime import date
from datetime import datetime

from openpyxl import load_workbook 

import customtkinter
import time
import os
import reportlab.rl_config 
import pyexcel as p
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import inch
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.units import inch,cm
reportlab.rl_config.warnOnMissingFontGlyphs = 0

from pdf2docx import parse 
from docx import Document 
