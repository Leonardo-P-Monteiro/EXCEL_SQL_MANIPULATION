import dotenv
import os
import pymysql
import openpyxl
from pathlib import Path

dotenv.load_dotenv()



# CONFIGURATION THE DATABASE.

connect = pymysql.connect(
    host = os.environ['SQL_HOST'],
    user= os.environ['SQL_USER'],
    password= os.environ['SQL_PASSWORD'],
    database= os.environ['SQL_DATABASE'],
)

cursor = connect.cursor()

# Creating the matrix table.
cursor.execute(
        'CREATE TABLE IF NOT EXISTS matrix '
        '('
        'id SMALLINT UNSIGNED PRIMARY KEY AUTO_INCREMENT, '
        'name_matter VARCHAR(250) NOT NULL '
        '); '               
               )

# Creating the subjects_table.
cursor.execute(
        'CREATE TABLE IF NOT EXISTS subjects '
        '('
        'id SMALLINT UNSIGNED PRIMARY KEY AUTO_INCREMENT, '
        'id_matrix SMALLINT UNSIGNED NOT NULL,'
        'name_subject VARCHAR(250) NOT NULL, '
        'conclusion_date DATE NULL, '
        'FOREIGN KEY (id_matrix) REFERENCES matrix(id) '
        'ON DELETE CASCADE ON UPDATE CASCADE'
        '); '               
        )



cursor.execute(
        'INSERT INTO matrix (name_matter) VALUES ("Test"); '
)

connect.commit()

cursor.close()
connect.close()

print('SQL Process conclude.')

# OPENING THE XLSX FILE

