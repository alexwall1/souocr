import sqlite3
from settings import DB_PATH


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("""CREATE TABLE IF NOT EXISTS file (id integer primary key autoincrement, md5_hash text,
    pdf_path text, xlsx_path text, pattern text, title text, subject text, author text, producer text, creator text, outline bool)""")

    c.execute("""CREATE TABLE IF NOT EXISTS page (id integer primary key autoincrement,
    file_id integer, box bool, body text, section text, page_number integer, 
    image_path text, x integer, y integer, width integer, height integer, FOREIGN KEY(file_id) references file(id))""")

    c.execute("""CREATE TABLE IF NOT EXISTS user (id integer primary key autoincrement,
    email text)""")

    conn.commit()
    return conn, c