import openpyxl
import pymysql


def capture_subjects(sheet_consult, _min_row, _max_row, colum_sql, id_matrix, cursor):
    """
    
    """
    for row in sheet_consult.iter_cols(min_col= 2, max_col= 2, min_row = 
            _min_row, max_row = _max_row, values_only= True):
        for value in row:
            if value == None:
                continue
            cursor.execute(
                f'INSERT IGNORE INTO subjects (id_matrix, {colum_sql}) VALUES (%s, %s)', (id_matrix, value)
            )