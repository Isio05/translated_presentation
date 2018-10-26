import os
import requests
from bs4 import BeautifulSoup as bs
import json
import zipfile
from shutil import copyfile, rmtree

def simple_open():
    f = open("test.txt", "r")
    f_lines = f.readlines()
    f.close()
    
    for line in f_lines:
        print(line)


def os_open_2(filename):
    fileDir = os.path.dirname(os.path.realpath(__file__))
    filepath = os.path.join(fileDir, filename)
    print(filepath)
    filehandle = open(filepath)
    f_lines = filehandle.readlines()
    for line in f_lines:
        print(line)
    filehandle.close()


def open_zip(archive_rel_path, num_of_slides):
    """Opens file contained in zip file without extraction"""
    slide_location = "ppt/slides/"
    slide_notation = "slideX.xml"
    
    # "archive" is to be open in read mode and is considered as source file
    archive = zipfile.ZipFile(os.path.join(os.path.dirname(os.path.realpath("__file__")), archive_rel_path), "r")
    # "arcive_2" will be an output file opened in write mode
    copyfile(os.path.join(os.path.dirname(__file__), archive_rel_path), os.path.join(os.path.dirname(__file__),
                                                                                     "".join([archive_rel_path[:archive_rel_path.find('.zip')],
                                                                                              "_translated_copy.zip"])))
    archive_2 = zipfile.ZipFile(os.path.join(os.path.dirname(__file__), "".join([archive_rel_path[:archive_rel_path.find('.zip')],
                                                                                 "_translated_copy.zip"])), "w")

    for item in archive.infolist():
        buffer = archive.read(item.filename)
        if not item.filename.startswith("ppt/slides/slide"):
            archive_2.writestr(item, buffer)
    
    translation = {}
    for slide in range(num_of_slides):
        # Slides are named according to convention in "slide_notation"
        # Open it directly from archive and bs starts to look for text to translate
        # After the translation the text is overwritten with translation
        # The end of the loop is to overwrite in archive_2 and start iteration on subsequent slide
        current_slide = slide_notation.replace("X", str(slide+1))
        current_slide_data = archive.read("".join([slide_location, current_slide]))
        xml_soup = bs(current_slide_data.decode("UTF-8"), 'lxml')

        for text in xml_soup.find_all("a:t"):
            translated = request_translation(text_input=text.string)
            translation[text.string] = translated['text'][0]

        current_slide_data_decoded = current_slide_data.decode("UTF-8")
        for item, definition in translation.items():
            current_slide_data_decoded = current_slide_data_decoded.replace(item, definition)
        current_slide_data_encoded = current_slide_data_decoded.encode("UTF-8")

        os.mkdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))
        f = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", current_slide), "wb")
        f.write(current_slide_data_encoded)
        f.close()
        archive_2.write(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp", current_slide),
                        "".join([slide_location, current_slide]))
        rmtree(os.path.join(os.path.dirname(os.path.realpath(__file__)), "temp"))

    archive_2.close()
    archive.close()
    
    return translation
    

def request_translation(text_input):
    url = "https://translate.yandex.net/api/v1.5/tr.json/translate"
    params = dict(key = "trnsl.1.1.20181021T090215Z.8bb8be7f52fe5b1e.10108f7b3312c6806795eb298f7960b1fd436295",
                  text = text_input.encode("UTF-8"),
                  lang = "pl-en")
    method = "GET"
    response = requests.request(url=url, method=method, params=params)
    content = json.loads(response.content)

    return content


##os_open_2(filename="test.txt")

##translated = request_translation(text_input="przykladowy tekst")
##print(translated['text'][0])

translated_pairs = open_zip(archive_rel_path="test.zip", num_of_slides=2)
[print(translated_pair) for translated_pair in translated_pairs.items()]
