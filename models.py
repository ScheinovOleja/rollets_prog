from pony.orm import Database, PrimaryKey, Required


db = Database()
db.bind(provider='sqlite', filename='database.sqlite', create_db=True)


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


def run():
    db.generate_mapping(create_tables=True)
