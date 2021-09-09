import datetime

from pony.orm import Database, PrimaryKey, Required

db = Database()
db.bind(provider='sqlite', filename='database.sqlite', create_db=True)


class Rollers(db.Entity):
    _table_ = 'rollers'
    id = PrimaryKey(int, auto=True)
    code = Required(int)
    name = Required(str)
    vendor = Required(str)
    color = Required(str)
    unit = Required(str)
    price = Required(float)
    availability = Required(int)


class DataNumDoc(db.Entity):
    _table_ = 'data_num_doc'
    id = PrimaryKey(int, auto=True)
    num_doc = Required(int)
    date = Required(str)


class AutomationData(db.Entity):
    _table_ = 'automation_data'
    id = PrimaryKey(int, auto=True)
    name_automatic = Required(str)
    price_automatic = Required(float)


db.generate_mapping(create_tables=True)
