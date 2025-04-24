import os
import win32print
import win32api
import win32com.client

def get_printers():
    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
    return printers

def print_pdf(filepath, printer_name=None):
    if printer_name is None:
        printer_name = win32print.GetDefaultPrinter()
    win32api.ShellExecute(0, "printto", filepath, f'"{printer_name}"', ".", 0)

def print_word(filepath, printer_name=None):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(filepath)
    doc.PrintOut(Printer=printer_name if printer_name else win32print.GetDefaultPrinter())
    doc.Close(False)
    word.Quit()

def print_file(filepath, printer_name=None):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        print_pdf(filepath, printer_name)
    elif ext in (".doc", ".docx"):
        print_word(filepath, printer_name)