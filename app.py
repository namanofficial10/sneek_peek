from flask import Flask, render_template, request, flash, redirect, url_for
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy
import os
import random
import math
from docx2pdf import convert
import aspose.words as aw
from PIL import ImageColor
import time



def doc_to_pdf(file_name, filepath = os.getcwd() + "\\static\\pdf_files"):

    try:
        convert(filepath + "\\" + file_name, keep_active=True, pbar_check=True)
        os.remove(filepath + "\\" + file_name)
        return True
    except:
        os.remove(filepath + "\\" + file_name)
        return False

upload_file_loc = r"C:\\Users\\DELL\\OneDrive\\Desktop\\PDF to ePUB sneek peek\\static\\pdf_files"
ALLOWED_EXTENSIONS = {'pdf','docx','doc'}
app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = "sqlite:///pdf_files.db"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
app.config['UPLOAD_FOLDER'] = upload_file_loc


def random_color():
    hexa_list = ['0', '1', '2', '3', '4', '5', '6',
                 '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']
    color = "#"
    while len(color) != 7:
        color += str(random.choice(hexa_list))
    

    return color

def isLightOrDark(rgbColor):
    [r,g,b]=rgbColor
    hsp = math.sqrt(0.299 * (r * r) + 0.587 * (g * g) + 0.114 * (b * b))
    if (hsp>127.5):
        return 'light'
    else:
        return 'dark'

def change_light_dark(current):
    if current == 'light':
        return "dark"
    else:
        return "light"


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


class confidential_naman(db.Model):
    sno = db.Column(db.Integer, primary_key=True)
    file_name = db.Column(db.String(100), nullable=False)
    ready_file_name = db.Column(db.String(100), nullable=False)

    def __repr__(self) -> str:
        return f"{self.sno}  {self.file_name}   {self.ready_file_name}"



def cvn(pdf_file_name):
    try:
        opath = r"C:\\Users\\DELL\\OneDrive\\Desktop\\PDF to ePUB sneek peek\\static\\pdf_files"
        file = pdf_file_name.replace(" ", "_")

        if file[-4:] != ".pdf":
            file = file + ".pdf"
        
        file_path = opath + f"\\{file}"

        doc = aw.Document(file_path)
        final_file_name = file[:-4]+".epub"
        doc.save(file_path[:-4]+".epub")
        print(file_path[:-4]+".epub")




    except Exception as e:
        print(e)
        final_file_name = pdf_file_name
    
    
    
    return final_file_name




@app.route("/")
def home(visibi="none"):
    color1 = random_color()
    color2 = random_color()
    color3 = random_color()
    
    color3_check = isLightOrDark(ImageColor.getcolor(color3, "RGB"))

    while True:
        if color3_check == "light":
            break;
        else:
            color3 = random_color()
            color3_check = isLightOrDark(ImageColor.getcolor(color3, "RGB"))    
    
    check_light_dark = isLightOrDark(ImageColor.getcolor(color1, "RGB"))

    check_light_dark_1 = change_light_dark(check_light_dark)

    return render_template("index.html", visi=visibi, color1 = color1, color2 = color2, light_or_dark = check_light_dark, light_or_dark_1 = check_light_dark_1, color3=color3)



@app.route("/db_starter")
def db_starter():
    entry = confidential_naman(
        sno=1, file_name="test.pdf", ready_file_name="test-converted.pdf")
    db.session.add(entry)
    db.session.commit()
    return render_template("index.html")


@app.route("/pdf_uploader", methods=["GET", "POST"])
def pdf_uploader():
    
    color1 = random_color()
    color2 = random_color()
    color3 = random_color()

    check_light_dark = isLightOrDark(ImageColor.getcolor(color1, "RGB"))
    check_light_dark_1 = change_light_dark(check_light_dark)

    if request.method == "POST":
        pdf_file = request.files['pdf_file']
        if pdf_file and allowed_file(pdf_file.filename):
            pdf_file.save(os.path.join(
                app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename)))
            if pdf_file.filename[-3:] == 'pdf':
                converted_file_name = cvn(pdf_file.filename)
                return render_template("download_converted_pdf.html", download_file=converted_file_name, color1 = color1, color2 = color2, light_or_dark = check_light_dark, light_or_dark_1 = check_light_dark_1)
            elif pdf_file.filename[-4:] == 'docx':
                word_file_name = pdf_file.filename.replace(" ", "_")
                time.sleep(5)
                check = doc_to_pdf(word_file_name)
                for i in range(3):
                    if check == True:
                        break
                    check = doc_to_pdf(word_file_name)
                # convert(os.getcwd() + "\\static\\pdf_files" + "\\" + pdf_file.filename, keep_active=True, pbar_check=True)
                try:
                    converted_file_name = cvn(word_file_name[:-5])
                    opath = r"C:\\Users\\DELL\\OneDrive\\Desktop\\PDF to ePUB sneek peek\\static\\pdf_files"
                    print(converted_file_name)
                    os.remove(opath + f"\\{word_file_name[:-5]}.pdf")
                    return render_template("download_converted_pdf.html", download_file=converted_file_name, color1 = color1, color2 = color2, light_or_dark = check_light_dark, light_or_dark_1 = check_light_dark_1)
                except:
                    return home(visibi="")
            else:
                return home(visibi="")


            # entry = confidential_naman(file_name = pdf_file, ready_file_name = converted_file_name)
            # db.session.add(entry)
            # db.session.commit()
    
            
        else:
            return home(visibi="")


if __name__ == "__main__":
    app.run(debug=True)