import csv

from PyQt5.QtWidgets import QFileDialog
from pony.orm import db_session

from models import AutomationData


@db_session
def delete_all_from_db(database):
    all_item = database.select()[:]
    for item in all_item:
        item.delete()


@db_session
def add_to_db(file, roll=None):
    encoding = 'cp1251' if roll else 'utf-8'
    with open(file, 'r', encoding=encoding) as file:
        reader = csv.reader(file, delimiter=';')
        delete_all_from_db(AutomationData)
        for row in reader:
            AutomationData(
                name_automatic=row[0],
                price_automatic=row[1],
            )
