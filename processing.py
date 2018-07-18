# coding=utf-8
import PyPDF2
import io
from wand.image import Image, Color
import cv2
import numpy
import imutils
from PIL import Image as PIL_Image
import pyocr.builders
import xlsxwriter
from subprocess import call
import os
from settings import PDF_FOLDER, XLSX_PATH, IMAGE_PATH, RESIZE_WIDTH, MIN_HEIGHT
from db import init_db
from flask import current_app
import re


def flatten_outline(pdf, document_info):
    result = []
    if isinstance(document_info, dict):
        result.append({'title': document_info['/Title'], 'page_number': pdf.getDestinationPageNumber(document_info)})
    elif isinstance(document_info, list):
        for section in document_info:
            result.extend(flatten_outline(pdf, section))
    else:
        raise Exception('Cannot read outline') # todo: fix error handling
    return result


def get_section_by_page(flat_outline, page_number):
    larger_than = [1 if el.get('page_number') <= page_number else 0 for el in flat_outline]
    index = sum(larger_than) - 2 # subtract one since its the last element, and one since its 0-indexed
    return flat_outline[index] if index > -1 else {'title': u'Avsnitt ej angivet', 'page_number': '-1'}


def preprocess_pdf(file_id):
    conn, c = init_db()
    values = (file_id,)
    results = c.execute("""SELECT pdf_path FROM file WHERE id = ?""", values).fetchone()
    pdf_path = results[0]

    pdf = PyPDF2.PdfFileReader(pdf_path)

    doc_info = pdf.getDocumentInfo()
    title = doc_info['/Title']
    subject = doc_info['/Subject']
    author = doc_info['/Author']
    producer = doc_info['/Producer']
    creator = doc_info['/Creator']
    outline = 1 if pdf.getOutlines() else 0

    values = (title, subject, author, producer, creator, outline, file_id)
    c.execute("""UPDATE file SET title = ?, subject = ?, author = ?, producer = ?, creator = ?, outline = ? WHERE id 
    = ?""", values)

    # if creator is Microsoft Word then preprocess pdf
    if u'Microsoft' in creator:
        if os.path.exists(pdf_path):
            abs_path_in = os.path.abspath(pdf_path)
            abs_path_out = os.path.join(PDF_FOLDER, '%s.pdf' % file_id)
            call(['gs', '-dBATCH', '-dNOPAUSE', '-q', '-sDEVICE=pdfwrite', '-sOutputFile=%s' % abs_path_out, abs_path_in])
            values = (abs_path_out, file_id)
            c.execute("""UPDATE file SET pdf_path = ? WHERE id = ?""", values)

    conn.commit()
    conn.close()

    return creator


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


def match_patter_in_pdf(file_id, pattern):
    conn, c = init_db()
    values = (file_id,)
    results = c.execute("""SELECT pdf_path, outline FROM file WHERE id = ?""", values).fetchone()
    pdf_path = results[0]
    has_outline = results[1]

    with io.open(pdf_path, 'rb') as f:
        pdf = PyPDF2.PdfFileReader(f)

        outline = pdf.getOutlines() if has_outline else None
        flat_outline = flatten_outline(pdf, outline) if outline else None

        current_app.logger.info('%s has %s pages' % (pdf_path, pdf.getNumPages()))

        for page_number in range(0, pdf.getNumPages()): # page_number is 0-indexed
            page = pdf.getPage(page_number)
            all_text = page.extractText()

            # todo: fix pattern matching
            if re.search(pattern.decode('utf-8'), all_text, re.I):
                if has_outline:
                    section = get_section_by_page(flat_outline, page_number)
                    values = (page_number, section.get('title'), file_id)
                    c.execute("""INSERT INTO page (page_number, section, file_id) values (?, ?, ?)""", values)
                else:
                    values = (page_number,)
                    c.execute("""INSERT INTO page (page_number) values (?)""", values)
                current_app.logger.info('Found pattern on page %s' % page_number)

    conn.commit()
    conn.close()


# todo: fix libpng warning
def generate_images_from_pdf(file_id):
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    values = (file_id,)
    results = c.execute("""SELECT p.id, p.page_number, f.pdf_path FROM page AS p
    INNER JOIN file AS f ON p.file_id = f.id WHERE f.id = ?""", values).fetchall()

    for page_id, page_number, pdf_path in results:
        with Image(filename='%s[%s]' % (pdf_path, page_number), resolution=300) as image:
            # convert from Wand Image object to numpy array
            image.background_color = Color('white')
            image.format = 'png'
            image.alpha_channel = 'background'
            image_buffer = numpy.asarray(bytearray(image.make_blob()), dtype=numpy.uint8)
            retval = cv2.imdecode(image_buffer, cv2.IMREAD_UNCHANGED) # todo: fix warnings

            # crop page
            current_app.logger.info('cropping page %s' % page_number)
            x, y, width, height, cropped_image = crop_image(retval)

            # prep results before saving to db

            # check if body in box
            if cropped_image.shape[0] > MIN_HEIGHT:
                image_path = '%s/%s_%s.png' % (IMAGE_PATH, file_id, page_number)
                cv2.imwrite(image_path, cropped_image)
                values = (image_path, True, x, y, width, height, page_id)
            else:
                values = (None, False, None, None, None, None, page_id)

            # save to db
            c.execute("""UPDATE page SET image_path = ?, box = ?, x = ?, y = ?, width = ?, height = ? WHERE id = ?""", values)
            current_app.logger.info('Saved image from page %s' % page_number)

    conn.commit()
    conn.close()


def ocr_images(file_id):
    # init db
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    values = (file_id,)

    # only OCR pages with boxes
    results = c.execute("""SELECT id, image_path FROM page WHERE box = 1 and file_id = ?""", values).fetchall()

    # init pyocr
    tools = pyocr.get_available_tools()
    tool = tools[0]
    langs = tool.get_available_languages()
    lang = langs[0]

    for page_id, image_path in results:
        current_app.logger.info('Running OCR on image id %s (%s)' % (page_id, image_path))

        # load image
        image = cv2.imread(image_path)
        text_image = PIL_Image.fromarray(image)

        # run ocr
        body = tool.image_to_string(text_image, lang=lang, builder=pyocr.builders.TextBuilder())

        # save results
        c.execute("""UPDATE page SET body = ? WHERE id = ?""", (body, page_id))

    conn.commit()
    conn.close()


def create_xlsx(file_id):
    conn, c = init_db()
    conn.row_factory = lambda cursor, row: row[0]
    values = (file_id,)

    results = c.execute("""SELECT title, subject, author, outline FROM file WHERE id = ?""", values).fetchone()
    title = results[0]
    subject = results[1]
    author = results[2]
    outline = results[3]

    results = c.execute("""SELECT body, page_number, section FROM page WHERE box = 1 AND file_id = ?""", values).fetchall()

    abs_path_out = os.path.join(XLSX_PATH, '%s.xlsx' % file_id)
    workbook = xlsxwriter.Workbook(abs_path_out)
    worksheet = workbook.add_worksheet()

    col = 0
    row = 0

    worksheet.write(row, col, title)
    worksheet.write(row, col + 1, subject)
    worksheet.write(row, col + 2, author)

    row += 2

    worksheet.write(row, col, 'Text')
    worksheet.write(row, col + 1, 'Sida')
    if outline:
        worksheet.write(row, col + 2, 'Avsnitt')
    else:
        worksheet.write(row, col + 2, '(Metadata i PDF med innehållsförteckning saknas)')

    row += 1

    for body, page_number, section in results:
        worksheet.write(row, col, body)
        worksheet.write(row, col + 1, page_number + 1)
        worksheet.write(row, col + 2, section)
        row += 1

    workbook.close()

    values = (abs_path_out, file_id)
    c.execute("""UPDATE file SET xlsx_path = ? WHERE id = ?""", values)
    conn.commit()
    conn.close()

    return abs_path_out


# if __name__ == '__main__':
#     from flask import Flask
#     app = Flask(__name__)
#     with app.app_context():
#         #preprocess_pdf(1)
#         #preprocess_pdf(2)
#         #match_patter_in_pdf(1, u'Utredningen föreslår')
#         #generate_images_from_pdf(1)
#         #ocr_images(1)
#         create_xlsx(1)