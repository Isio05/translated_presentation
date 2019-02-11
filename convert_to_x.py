from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants


def save_as_docx(path, word):
    print(path)
    # If file is corrupted or locked it will be skipped
    try:
        doc = word.Documents.Open(path)
        doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        doc.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        doc.Close()
    except:
        print("Incorrect file")


def save_as_xlsx(path, excel):
    print(path)
    # If file is corrupted or locked it will be skipped
    try:
        wb = excel.Workbooks.Open(path)

        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.xlsx', new_file_abs)

        wb.SaveAs(new_file_abs, FileFormat=constants.xlOpenXMLWorkbook	)
        wb.Close()
    except:
        print("Incorrect file")


def save_as_pptx(path, powerpoint):
    # If file is corrupted or locked it will be skipped
    print(path)
    try:
        pres = powerpoint.Presentations.Open(path)

        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.pptx', new_file_abs)

        pres.SaveAs(new_file_abs, FileFormat=constants.ppSaveAsOpenXMLPresentation)
        pres.Close()
    except:
        print("Incorrect file")


def convert_rtf_doc(path):
    # Very simple conversion of .rtf file to .doc so it can be used by save_as_docx function
    print(path)
    if len(re.findall(".rtf$", path)) == 1:
        os.rename(path, re.sub(r'\.rtf$', '.doc', path))
        path = re.sub(r'\.rtf$', '.doc', path)


def change_all_to_x():
    path = input("Set full path to folder: ")

    # Sample correct final path:
    # C:\\Users\\user\\PycharmProjects\\translate_presentation\\subfolder\\**\*.xls

    # Convert .rtf files to .doc part
    rtf_files = glob(path + "\\**\*.rtf", recursive=True)
    for rtf in rtf_files:
        convert_rtf_doc(rtf)

    # Convert .doc to .docx part
    # Create interface
    word = win32.gencache.EnsureDispatch('Word.Application')
    # Create list of paths to files
    doc_files = glob(path + "\\**\*.doc", recursive=True)
    for doc in doc_files:
        save_as_docx(doc, word)
    word.Quit()

    # Convert .ppt to .pptx part
    # Create interface
    powerpoint = win32.gencache.EnsureDispatch('Powerpoint.Application')
    # Create list of paths to files
    ppt_files = glob(path + "\\**\*.ppt", recursive=True)
    for ppt in ppt_files:
        save_as_pptx(ppt, powerpoint)
    powerpoint.Quit()

    # Convert .xls to .xlsx part
    # Create interface
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # Create list of paths to files
    xls_files = glob(path + "\\**\*.xls", recursive=True)
    for xls in xls_files:
        save_as_xlsx(xls, excel)
    excel.Application.Quit()


change_all_to_x()
