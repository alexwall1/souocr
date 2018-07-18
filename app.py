from flask import Flask, render_template, request, redirect, flash
from settings import UPLOAD_FOLDER, SECRET_KEY, CELERY_BROKER_URL, CELERY_RESULT_BACKEND, MAIL_USERNAME, MAIL_PASSWORD, MAIL_SERVER, MAIL_DEFAULT_SENDER
from db import init_db
from flask_mail import Mail

import tasks
import hashlib
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = SECRET_KEY
app.config.update(CELERY_BROKER_URL=CELERY_BROKER_URL,
                  CELERY_RESULT_BACKEND=CELERY_RESULT_BACKEND)

app.config.update(
    DEBUG=True,
    MAIL_SERVER=MAIL_SERVER,
    MAIL_PORT=587,
    MAIL_USE_SSL=False,
    MAIL_USE_TLS=True,
    MAIL_USERNAME=MAIL_USERNAME,
    MAIL_PASSWORD=MAIL_PASSWORD,
    MAIL_DEFAULT_SENDER=MAIL_DEFAULT_SENDER
)
mail = Mail(app)


def md5_for_file(f, block_size=2**20):
    data = True
    md5 = hashlib.md5()
    while data:
        data = f.read(block_size)
        if data:
            md5.update(data)
    f.seek(0)
    return md5.hexdigest()


@app.route('/')
def home():
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    files = c.execute("""SELECT pdf_path, xlsx_path, title, subject, author FROM file""").fetchall()
    return render_template('home.html', files=files)


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        email = request.form['email']
        conn, c = init_db()
        results = c.execute("""SELECT count(*) FROM user WHERE email == ?""", (email,)).fetchone()
        if results[0] < 1:
            flash('Invalid e-mail address.')
            return redirect(request.url)
        # todo: check e-mail
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # todo: make content-type check smarter
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        # todo: fix error-handling
        if file and file.content_type == 'application/pdf':
            md5_hash = md5_for_file(file.stream)
            values = (md5_hash,)
            results = c.execute("""SELECT COUNT(*) FROM file WHERE md5_hash == ?""", values).fetchone()
            if results[0] < 1:
                pdf_path = os.path.join(UPLOAD_FOLDER, '%s.pdf' % md5_hash)
                values = (md5_hash, pdf_path)
                c.execute("""INSERT INTO file (md5_hash, pdf_path) values (?, ?)""", values)
                conn.commit()
                flash('Upload successful!')
                file.save(pdf_path)

                # start asynchronous processing
                file_id = c.lastrowid
                tasks.process_file.delay(email, file_id)
                return render_template('success.html', email=email)
            else:
                flash('File (%s) already exists, aborting.' % md5_hash)
                return render_template('error.html')
        else:
            flash('File upload failed')
            return redirect(request.url)

    # if method GET render upload form
    return render_template('upload.html')


if __name__ == '__main__':
    app.run(debug=True)