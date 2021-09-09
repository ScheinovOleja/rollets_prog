import csv

from PyQt5.QtWidgets import QFileDialog
from pony.orm import db_session

from models import Rollers


@db_session
def delete_all_from_db():
    all_roll = Rollers.select().order_by(Rollers.id)
    for roll in all_roll:
        roll.delete()


@db_session
def add_to_db(file):
    delete_all_from_db()
    with open(file, 'r', encoding='windows-1251') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            price = float(row[5].replace(',', '.'))
            Rollers(
                code=row[0],
                name=row[1],
                vendor=row[2],
                color=row[3],
                unit=row[4],
                price=price,
                availability=row[6],
            )
