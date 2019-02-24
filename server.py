from open_it_aws import TranslatePresentation, TranslateWorkbook, TranslateDocument
from flask import Flask, render_template, redirect, url_for, session, request, send_from_directory, make_response
from tempfile import mkdtemp
from utils import LANGUAGE_PAIRS, CODE_PAIRS, ALLOWED_EXTENSIONS
from shared_variables import SECRET_KEY, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY
from werkzeug.utils import secure_filename
from datetime import datetime
import numpy as np
import boto3
import zipfile
import json
import random
import string
import os

app = Flask(__name__, static_folder="Static", template_folder="Templates")
# Configure session to use filesystem (instead of signed cookies)
app.config["SESSION_FILE_DIR"] = mkdtemp()
app.config["SESSION_PERMANENT"] = True
app.config["SESSION_TYPE"] = "filesystem"
# Ensure templates are auto-reloaded
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["UPLOAD_FOLDER"] = os.path.dirname(__file__)
app.secret_key = SECRET_KEY


@app.route("/")
def index():
    return redirect(url_for("translate", input_l="pl", output_l="en"))


@app.route("/error")
def error():
    return json.loads(session["message"])


@app.route("/translate/<input_l>/<output_l>", methods=["GET", "POST"])
def translate(input_l, output_l):
    if request.method == "GET":
        return render_template("translate.html", language_pairs=LANGUAGE_PAIRS.values(),
                               input_l=LANGUAGE_PAIRS[input_l], output_l=LANGUAGE_PAIRS[output_l])
    elif request.method == "POST":
        new_input_l = CODE_PAIRS[request.form.get("input_l")]
        new_output_l = CODE_PAIRS[request.form.get("output_l")]

        # Using class methods change used language for each subsequent translation
        # Changes in languages will be inherited by other classes (Doc and Xls translation class inherit
        # translation method and its settings from the Presentation class)
        TranslatePresentation.change_input_language(new_input_l=new_input_l)
        TranslatePresentation.change_ouput_language(new_output_l=new_output_l)

        files = request.files.getlist('files')

        # Check if user sent any file
        if files == "":
            session["message"] = json.dumps("No file sent")
            return redirect(url_for("error"))

        # Create temp folder for files and translated copies
        # Pass temp folder by the class function, as Doc and Xls translation classes inherit from the Ppt
        # there is only need to change it in the super class
        temp_folder = ''.join(random.choices(string.ascii_letters + string.digits, k=16))
        os.mkdir(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder))
        TranslatePresentation.change_temp_folder(new_temp_folder=temp_folder)

        # Write files to the temporary folder
        for file in files:
            # Check if extension is correct
            if "." not in file.filename or file.filename.rsplit(".", 1)[1].lower() not in ALLOWED_EXTENSIONS:
                session["message"] = json.dumps("Wrong extension of sent files")
                return redirect(url_for("error"))
            # Write secured filename to the variable
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder, filename))

        # Open archive where translated files will be saved
        translated_files = zipfile.ZipFile(
            os.path.join(app.config['UPLOAD_FOLDER'], temp_folder, "translated_files.zip"), "x")

        # Iterate through every file in the folder and translate it choosing appropriate class
        for f in [file for file in os.listdir(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder))
                  if os.path.isfile(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder, file)) and
                     os.path.splitext(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder, file))[1] != ".zip"]:
            file_type = os.path.splitext(f)[1]
            if file_type == ".docx":
                translate = TranslateDocument(file_to_translate=f)
                translated_file_coords = translate.main()
            elif file_type == ".pptx":
                translate = TranslatePresentation(file_to_translate=f)
                translated_file_coords = translate.main()
            elif file_type == ".xlsx":
                translate = TranslateWorkbook(file_to_translate=f)
                translated_file_coords = translate.main()
            else:
                raise RuntimeError("Wrong file extension")
            # Remove source file for translation
            os.remove(os.path.join(app.config["UPLOAD_FOLDER"], temp_folder, f))
            # Write translated file to archive
            translated_files.write(translated_file_coords['translated_file_path'],
                                   arcname=translated_file_coords['translated_file'])
            # Remove translated file as it is contained within the archive
            os.remove(
                os.path.join(app.config["UPLOAD_FOLDER"], temp_folder, translated_file_coords['translated_file_path']))

        # Upload translated files to S3 bucket
        # s3 = boto3.client(service_name='s3', region_name='us-east-1', use_ssl=True,
        #                   aws_access_key_id=AWS_ACCESS_KEY_ID,
        #                   aws_secret_access_key=AWS_SECRET_ACCESS_KEY
        #                   )

        # Configure resource
        s3 = boto3.resource(service_name='s3', region_name='us-east-1', use_ssl=True,
                            aws_access_key_id=AWS_ACCESS_KEY_ID,
                            aws_secret_access_key=AWS_SECRET_ACCESS_KEY
                            )

        # Create random user name for not logged users - they will have their own folders on S3
        try:
            user_name = session['user']
        except KeyError:
            random_string = random.choices(string.ascii_letters + string.digits, k=16)
            random_string.insert(0, "anon_")
            user_name = ''.join(random_string)
            session['user'] = user_name

        # Use resource to call Object object
        s3_object = s3.Object(bucket_name="translatedfiles",
                              key=str(user_name + "/" + temp_folder + "/" + "translated_files.zip"))
        s3_object.upload_file(
            os.path.join(app.config['UPLOAD_FOLDER'], temp_folder, "translated_files.zip"))

        cookie = request.cookies.get("translated_files_list")
        print(cookie)

        # Cookies have following format: <temp_folder for translation_1>,<time of translation_1>,
        # <temp_folder for translation_2>,<time of translation_2> etc.
        # The first parameter allows to retrieve files from S3, time is used for listing
        if not cookie:
            res = make_response(redirect(url_for("translate", input_l=new_input_l, output_l=new_output_l)))
            res.set_cookie("translated_files_list", temp_folder + "," + str(datetime.now()), 60 * 60 * 24 * 30)
        else:
            cookie += "," + temp_folder + "," + str(datetime.now())
            print(cookie)
            res = make_response(redirect(url_for("translate", input_l=new_input_l, output_l=new_output_l)))
            res.set_cookie("translated_files_list", cookie, 60 * 60 * 24 * 30)

        # return redirect(url_for("translate", input_l=new_input_l, output_l=new_output_l))
        return res

    else:
        session["message"] = json.dumps("Wrong method")
        return redirect(url_for("error"))


@app.route("/translated-files")
def translated_files():
    # Gather list of files and date from the cookie file
    cookie = request.cookies.get("translated_files_list").split(",")
    # Convert list to array that can be easily used
    files_array = np.array(cookie, dtype=str)
    files_array = files_array.reshape((-1,2))
    files_array = np.hstack((np.arange(files_array.shape[0]).reshape((files_array.shape[0],1)), files_array))

    return render_template("translated_files.html",
                           t_names=files_array[:, 0].tolist(),
                           t_dates=files_array[:, 2].tolist(),
                           length = len(files_array[:, 0].tolist()))


app.run(port=4544, debug=True)
