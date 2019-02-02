import os
import requests
from bs4 import BeautifulSoup as bs
import json
import zipfile
from shutil import copyfile, rmtree
from shared_variables import API_KEY


class TranslatePresentation:
    def __init__(self, file_to_translate):
        self.old_extension = None
        self.file_ready_to_translate = None
        self.file_to_translate = None
        self.user_num_of_slides = None
        self.file_to_translate = file_to_translate
        self.num_of_slides = 0

    def open_zip(self):
        """Opens file contained in zip file without extraction"""
        # Relative paths used to navigate within xml file
        slide_location = "ppt/slides/"
        slide_notation = "slideX.xml"

        # "archive" is to be open in read mode and is considered as source file
        archive = zipfile.ZipFile(
            os.path.join(os.path.dirname(os.path.realpath("__file__")), self.file_ready_to_translate), "r")
        # "archive_2" will be an output file opened in write mode
        copyfile(os.path.join(os.path.dirname(__file__), self.file_ready_to_translate),
                 os.path.join(os.path.dirname(__file__),
                              "".join([self.file_ready_to_translate[
                                       :self.file_ready_to_translate.find(
                                           '.zip')],
                                       "_translated_copy.zip"])))
        archive_2 = zipfile.ZipFile(
            os.path.join(os.path.dirname(__file__),
                         "".join([self.file_ready_to_translate[:self.file_ready_to_translate.find('.zip')],
                                  "_translated_copy.zip"])), "w")

        # Rewrite each file separately from the source archive to containing translation
        for item in archive.infolist():
            buffer = archive.read(item.filename)
            if not item.filename.startswith("ppt/slides/slide"):
                archive_2.writestr(item, buffer)
            else:
                self.num_of_slides += 1

        # To write into archive, the source file must exist
        # The "temp" folder will contain ready to write xmls with translations
        # After the operation the folder temp is removed
        os.mkdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))

        translation = {}
        for slide in range(self.num_of_slides):
            # Slides are named according to convention in "slide_notation"
            # Open it directly from archive and bs starts to look for text to translate
            # After the translation the text is overwritten with translation
            # The end of the loop is to write in archive_2 and start iteration on subsequent slide
            current_slide = slide_notation.replace("X", str(slide + 1))
            current_slide_data = archive.read("".join([slide_location, current_slide]))
            xml_soup = bs(current_slide_data.decode("UTF-8"), 'lxml')

            # Text on each slide is surrounded by "a:t"
            # Loop iterates through each mark in the xml file, sends it to API and saves translation to dictionary
            for text in xml_soup.find_all("a:t"):
                translated = self.request_translation(text_input=text.string)
                translation[text.string] = translated['text'][0]

            # The source slide is unpacked into simple string
            # Using dictionary that contains translations and sources, text will be replaced in the string
            # After the operation, string is encoded to basic format
            current_slide_data_decoded = current_slide_data.decode("UTF-8")
            for item, definition in translation.items():
                current_slide_data_decoded = current_slide_data_decoded.replace("<a:t>" + item + "</a:t>",
                                                                                "<a:t>" + definition + "</a:t>")
            current_slide_data_encoded = current_slide_data_decoded.encode("UTF-8")

            # Using created temp folder, create there xml file containing translation
            # Subsequently, write (wb) that file to translation archive
            f = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", current_slide), "wb")
            f.write(current_slide_data_encoded)
            f.close()
            archive_2.write(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", current_slide),
                            "".join([slide_location, current_slide]))

        # Remove temp folder that contain xmls with translation
        rmtree(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))

        archive_2.close()
        archive.close()

        return translation

    @staticmethod
    def request_translation(text_input):
        url = "https://translate.yandex.net/api/v1.5/tr.json/translate"
        params = dict(key=API_KEY,
                      text=text_input.encode("UTF-8"),
                      lang="pl-en")
        method = "GET"
        response = requests.request(url=url, method=method, params=params)
        content = json.loads(response.content)

        return content

    def convert_file_ext(self):
        """Changes the extension of file, from ppt(x) to zip and backwards"""
        # Function will try to recognize current extension and choose conversion direction
        archive_abs_path = os.path.join(os.path.dirname(__file__), self.file_to_translate)
        # Split archive path into path itself and extension
        archive_split = os.path.splitext(archive_abs_path)
        if archive_split[1] in (".pptx", ".docx", ".xlsx"):
            # Rename using original absolute path and that path with modified extension
            os.rename(archive_abs_path, archive_split[0] + ".zip")
            self.file_ready_to_translate = str(archive_split[0].split("\\")[1]) + ".zip"
            self.old_extension = archive_split[1]
            # Old extension is returned for usage in backwards conversion
        elif archive_split[1] == ".zip":
            # Rename using original absolute path and that path with modified extension
            os.rename(archive_abs_path, archive_split[0] + self.old_extension)
        else:
            print("Wrong extension of provided file.")

    def main(self):
        # Archive relative path - currently searches the catalog of script location
        # Change file extension to .zip and write to variable its changed name
        self.convert_file_ext()
        # Perform translation and print out the translated texts
        translated_pairs = self.open_zip()
        [print(translated_pair) for translated_pair in translated_pairs.items()]
        # Change extensions of original and translated file back to ".ppt(x)"
        self.file_to_translate = self.file_ready_to_translate
        self.convert_file_ext()
        self.file_to_translate = self.file_to_translate.replace(".zip", "_translated_copy.zip")
        self.convert_file_ext()
        print("Done")


class TranslateDocument(TranslatePresentation):
    def open_zip(self):
        """Opens file contained in zip file without extraction"""
        # Relative paths used to navigate within xml file
        contents_file_location = "word/"
        contents_file = "document.xml"

        # "archive" is to be open in read mode and is considered as source file
        archive = zipfile.ZipFile(
            os.path.join(os.path.dirname(os.path.realpath("__file__")), self.file_ready_to_translate), "r")
        # "archive_2" will be an output file opened in write mode
        copyfile(os.path.join(os.path.dirname(__file__), self.file_ready_to_translate),
                 os.path.join(os.path.dirname(__file__),
                              "".join([self.file_ready_to_translate[
                                       :self.file_ready_to_translate.find(
                                           '.zip')],
                                       "_translated_copy.zip"])))
        archive_2 = zipfile.ZipFile(
            os.path.join(os.path.dirname(__file__),
                         "".join([self.file_ready_to_translate[:self.file_ready_to_translate.find('.zip')],
                                  "_translated_copy.zip"])), "w")

        # Rewrite each file separately from the source archive to containing translation
        for item in archive.infolist():
            buffer = archive.read(item.filename)
            if not item.filename.startswith("word/document"):
                archive_2.writestr(item, buffer)

        translation = {}
        # Open it directly from archive and bs starts to look for text to translate
        # After the translation the text is overwritten with translation
        # The end of the loop is to write into archive_2
        contents_file_rel_path = contents_file_location + contents_file
        current_slide_data = archive.read(contents_file_rel_path)
        xml_soup = bs(current_slide_data.decode("UTF-8"), 'lxml')

        # Text on each slide is surrounded by "w:t"
        # Loop iterates through each mark in the xml file, sends it to API and saves translation to dictionary
        for text in xml_soup.find_all("w:t"):
            translated = super().request_translation(text_input=text.string)
            translation[text.string] = translated['text'][0]

        # The source slide is unpacked into simple string
        # Using dictionary that contains translations and sources, text will be replaced in the string
        # After the operation, string is encoded to basic format
        current_slide_data_decoded = current_slide_data.decode("UTF-8")
        for item, definition in translation.items():
            current_slide_data_decoded = current_slide_data_decoded.replace("<w:t>" + item + "</w:t>",
                                                                            "<w:t>" + definition + "</w:t>")
        current_slide_data_encoded = current_slide_data_decoded.encode("UTF-8")

        # To write into archive, the source file must exist
        # The "temp" folder will contain ready to write xmls with translations
        # After the operation the folder temp is removed
        # Using created temp folder, create there xml file containing translation
        # Subsequently, write (wb) that file to translation archive
        os.mkdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))
        f = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", contents_file), "wb")
        f.write(current_slide_data_encoded)
        f.close()
        archive_2.write(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", contents_file),
                        contents_file_rel_path)
        rmtree(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))

        archive_2.close()
        archive.close()

        return translation


class TranslateWorkbook(TranslatePresentation):
    def open_zip(self):
        """Opens file contained in zip file without extraction"""
        # Relative paths used to navigate within xml file
        contents_file_location = "xl/"
        contents_file = "sharedStrings.xml"

        # "archive" is to be open in read mode and is considered as source file
        archive = zipfile.ZipFile(
            os.path.join(os.path.dirname(os.path.realpath("__file__")), self.file_ready_to_translate), "r")
        # "archive_2" will be an output file opened in write mode
        copyfile(os.path.join(os.path.dirname(__file__), self.file_ready_to_translate),
                 os.path.join(os.path.dirname(__file__),
                              "".join([self.file_ready_to_translate[
                                       :self.file_ready_to_translate.find(
                                           '.zip')],
                                       "_translated_copy.zip"])))
        archive_2 = zipfile.ZipFile(
            os.path.join(os.path.dirname(__file__),
                         "".join([self.file_ready_to_translate[:self.file_ready_to_translate.find('.zip')],
                                  "_translated_copy.zip"])), "w")

        # Rewrite each file separately from the source archive to containing translation
        for item in archive.infolist():
            buffer = archive.read(item.filename)
            if not item.filename.startswith("xl/sharedStrings"):
                archive_2.writestr(item, buffer)

        translation = {}
        # Open it directly from archive and bs starts to look for text to translate
        # After the translation the text is overwritten with translation
        # The end of the loop is to write into archive_2
        contents_file_rel_path = contents_file_location + contents_file
        current_slide_data = archive.read(contents_file_rel_path)
        xml_soup = bs(current_slide_data.decode("UTF-8"), 'lxml')

        # Text on each slide is surrounded by "w:t"
        # Loop iterates through each mark in the xml file, sends it to API and saves translation to dictionary
        for text in xml_soup.find_all("t"):
            translated = super().request_translation(text_input=text.string)
            translation[text.string] = translated['text'][0]

        # The source slide is unpacked into simple string
        # Using dictionary that contains translations and sources, text will be replaced in the string
        # After the operation, string is encoded to basic format
        current_slide_data_decoded = current_slide_data.decode("UTF-8")
        for item, definition in translation.items():
            current_slide_data_decoded = current_slide_data_decoded.replace("<t>" + item + "</t>",
                                                                            "<t>" + definition + "</t>")
        current_slide_data_encoded = current_slide_data_decoded.encode("UTF-8")

        # To write into archive, the source file must exist
        # The "temp" folder will contain ready to write xmls with translations
        # After the operation the folder temp is removed
        # Using created temp folder, create there xml file containing translation
        # Subsequently, write (wb) that file to translation archive
        os.mkdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))
        f = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", contents_file), "wb")
        f.write(current_slide_data_encoded)
        f.close()
        archive_2.write(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", contents_file),
                        contents_file_rel_path)
        rmtree(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))

        archive_2.close()
        archive.close()

        return translation


def menu():
    while True:
        file = input("Type in file type with extension or 'exit': ")
        if file == "exit":
            break
        file_type = os.path.splitext(file)[1]
        if file_type == ".docx":
            translate = TranslateDocument(file_to_translate=file)
            translate.main()
        elif file_type == ".pptx":
            translate = TranslatePresentation(file_to_translate=file)
            translate.main()
        elif file_type == ".xlsx":
            translate = TranslateWorkbook(file_to_translate=file)
            translate.main()
        else:
            print("Wrong file extension")


menu()
