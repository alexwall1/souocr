# -*- coding: utf-8 -*-

import PyPDF2
import io
from wand.image import Image, Color
import cv2
import numpy
import imutils
from PIL import Image as PIL_Image
import pyocr.builders
import sqlite3
from hashlib import md5
import xlsxwriter
from subprocess import call
import os
from settings import DB_PATH, PDF_PATH, XLSX_PATH, IMAGE_PATH, RESIZE_WIDTH, MIN_HEIGHT


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS result (id integer primary key autoincrement, md5_hash text, pdf_path text, box bool, body text, page integer, 
    image_path text, x integer, y integer, width integer, height integer)""")
    conn.commit()
    return conn, c


def prepare_pdf(pdf_path):
    if os.path.exists(pdf_path):
        abs_path_in = os.path.abspath(pdf_path)
        md5_hash = md5(pdf_path).hexdigest()
        abs_path_out = os.path.join(PDF_PATH, '%s.pdf' % md5_hash)
        call(['gs', '-dBATCH', '-dNOPAUSE', '-q', '-sDEVICE=pdfwrite', '-sOutputFile=%s' % abs_path_out, abs_path_in])
        return abs_path_out


def crop_image(image):
    ratio = float(RESIZE_WIDTH) / image.shape[1]

    # resize and transform
    resized = imutils.resize(image, width=RESIZE_WIDTH)

    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    ret, thresh = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)
    negative = cv2.bitwise_not(thresh)

    # find contours
    contours = cv2.findContours(negative, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if imutils.is_cv2() else contours[1]

    # find rectangle
    max_perimeter = 0
    max_contour = None
    for index, contour in enumerate(contours):
        perimeter = cv2.arcLength(contour, False)
        if perimeter > max_perimeter:
            max_perimeter = perimeter
            max_contour = contour

    approx = cv2.approxPolyDP(max_contour, 0.04 * max_perimeter, False) # todo: test if necessary
    (x, y, width, height) = cv2.boundingRect(approx)

    # re-scale
    x_p = int(1 / ratio * x)
    y_p = int(1 / ratio * y)
    height_p = int(1 / ratio * height)
    width_p = int(1 / ratio * width)

    return x_p, y_p, width_p, height_p, image[y_p:y_p + height_p, x_p:x_p + width_p]


def match_patter_in_pdf(pdf_path, pattern):
    conn, c = init_db()
    md5_hash = md5(pdf_path).hexdigest()

    with io.open(pdf_path, 'rb') as f:
        pdf = PyPDF2.PdfFileReader(f)

        print '%s has %s pages' % (pdf_path, pdf.getNumPages())

        for page_number in range(0, pdf.getNumPages()):
            page = pdf.getPage(page_number)
            all_text = page.extractText()
            if pattern in all_text:
                values = (page_number, md5_hash)
                c.execute("""INSERT INTO result (page, md5_hash) values (?, ?)""", values)
                print 'Found pattern on page %s' % page_number
    conn.commit()
    conn.close()
    return md5_hash


def generate_images_from_pdf(pdf_path):
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]

    md5_hash = md5(pdf_path).hexdigest()

    values = (md5_hash,)
    results = c.execute("""SELECT id,page FROM result WHERE md5_hash = ?""", values).fetchall()

    for _id, page_number in results:
        with Image(filename='%s[%s]' % (pdf_path, page_number), resolution=300) as image:
            # convert from Wand Image object to numpy array
            image.background_color = Color('white')
            image.format = 'png'
            image.alpha_channel = 'background'
            image_buffer = numpy.asarray(bytearray(image.make_blob()), dtype=numpy.uint8)
            retval = cv2.imdecode(image_buffer, cv2.IMREAD_UNCHANGED)

            # crop page
            print 'cropping page %s' % page_number
            x, y, width, height, cropped_image = crop_image(retval)

            # prep results before saving to db
            if cropped_image.shape[0] > MIN_HEIGHT:
                image_path = '%s/%s_%s.png' % (IMAGE_PATH, md5_hash, page_number)
                cv2.imwrite(image_path, cropped_image)
                values = (image_path, True, x, y, width, height, _id)
            else:
                values = (None, False, None, None, None, None, _id)

            # save to db
            c.execute("""UPDATE result SET image_path = ?, box = ?, x = ?, y = ?, width = ?, height = ? WHERE id = ?""", values)
            print 'Saved image box from page %s' % page_number
    conn.commit()
    conn.close()
    return md5_hash


def ocr_images(md5_hash):
    # init db
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    values = (md5_hash,)
    results = c.execute("""SELECT id, image_path FROM result WHERE box = 1 and md5_hash = ?""", values).fetchall()

    # init pyocr
    tools = pyocr.get_available_tools()
    tool = tools[0]
    langs = tool.get_available_languages()
    lang = langs[0]

    for _id, image_path in results:
        print 'Running OCR on image id %s (%s)' % (_id, image_path)

        # load image
        image = cv2.imread(image_path)
        text_image = PIL_Image.fromarray(image)

        # run ocr
        body = tool.image_to_string(text_image, lang=lang, builder=pyocr.builders.TextBuilder())

        # save results
        c.execute("""UPDATE result SET body = ? WHERE id = ?""", (body, _id))
    conn.commit()
    conn.close()


def write_xlsx(md5_hash):
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    values = (md5_hash,)
    results = c.execute("""SELECT body, page FROM result WHERE box = 1 AND md5_hash = ?""", values).fetchall()

    abs_path_out = os.path.join(XLSX_PATH, '%s.xlsx' % md5_hash)
    workbook = xlsxwriter.Workbook(abs_path_out)
    worksheet = workbook.add_worksheet()

    col = 0
    row = 0

    worksheet.write(row, col, 'Text')
    worksheet.write(row, col + 1, 'Sida')

    row += 1

    for body, page_number in results:
        worksheet.write(row, col, body)
        worksheet.write(row, col + 1, page_number)
        row += 1

    workbook.close()
    conn.close()


if __name__ == '__main__':
    pattern = u'Utredningen föreslår'

    pdf_path = prepare_pdf('./pdfs/df.pdf')
    md5_hash = match_patter_in_pdf(pdf_path, pattern)
    generate_images_from_pdf(pdf_path)
    ocr_images(md5_hash)
    write_xlsx(md5_hash)

