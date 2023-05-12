import os
from flask import Flask, render_template, request, flash, redirect, send_file
import config
from app.controller.fgis import FileXLSX, VerificationCert


app = Flask(__name__)

app.config.from_object(config.Config)


@app.get('/')
def fgis():
    return render_template('fgis.html')


@app.post('/')
def view():
    file = request.files['file']
    if file.filename == '':
        flash('Не выбран файл')
        return redirect(request.url)

    xlsx_to_upload = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(xlsx_to_upload)
    xlsx = FileXLSX(xlsx_to_upload)
    cert = VerificationCert(xlsx.data)
    cert.get_sert()
    download_file = cert.add_to_zip()

    return send_file(os.path.join(os.path.abspath(app.config['DOWNLOAD_FOLDER']), download_file), as_attachment=True)
