import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug import secure_filename
from flattenizer import flattenizer
import pandas as pd
from os.path import join

# Initialize the Flask application
app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = set(['xlsx'])

# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    return render_template('flattenizer.html')

# Route that will process the file upload
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(join(app.config['UPLOAD_FOLDER'], filename))
        flattenizer(join(app.config['UPLOAD_FOLDER'], filename), join(app.config['UPLOAD_FOLDER']))

        return redirect(url_for('uploaded_file',
                                filename='FLAT_' + filename))


        #return "File flattenized!!!"
    else:
        return render_template('nofile.html')

# This figure out why this works down there
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)

if __name__ == '__main__':
    app.run(
        host="0.0.0.0",
        port=int("9000"),
        debug=True
    )
