# coding=utf-8
from db import init_db
from processing import preprocess_pdf, match_patter_in_pdf, generate_images_from_pdf, ocr_images, create_xlsx
from flask_mail import Message
from app import app, mail
from celery import Celery


def make_celery(app):
    celery = Celery(app.import_name,
                    backend=app.config['CELERY_RESULT_BACKEND'],
                    broker=app.config['CELERY_BROKER_URL'])
    celery.conf.update(app.config)
    TaskBase = celery.Task

    class ContextTask(TaskBase):
        abstract = True

        def __call__(self, *args, **kwargs):
            with app.app_context():
                return TaskBase.__call__(self, *args, **kwargs)

    celery.Task = ContextTask
    return celery


celery = make_celery(app)


@celery.task(name='tasks.process_file')
def process_file(email, file_id):
    conn, c = init_db()

    results = c.execute("""SELECT id FROM file WHERE id = ?""", (file_id,)).fetchone()
    file_id = results[0]
    pattern = 'Utredningen (föreslår|bedömer)'

    creator = preprocess_pdf(file_id)
    app.logger.info('Pre-processing done of file %s.' % file_id)

    message = Message("Hello", recipients=[email])

    if u'Microsoft' in creator:
        match_patter_in_pdf(file_id, pattern)
        app.logger.info('Done matching pattern in file %s.' % file_id)

        generate_images_from_pdf(file_id)
        app.logger.info('Done generating images from file %s' % file_id)

        ocr_images(file_id)
        app.logger.info('Done OCR\'ing images from pdf %s' % file_id)

        xlsx_path = create_xlsx(file_id)
        app.logger.info('Created XLSX file %s' % xlsx_path)

        message.body = u'Please find attached the result XLSX file.'

        with app.open_resource(xlsx_path) as f:
            message.attach(xlsx_path.split("/")[-1], 'application/pdf', f.read())
    else:
        message.body = u'Only native PDF files created by Microsoft Word are supported. The uploaded PDF was created ' \
                       u'by: %s.' % creator

    mail.send(message)
    app.logger.info('E-mail sent to %s' % email)
