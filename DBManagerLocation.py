import sqlite3
from DBManager import DBManager

class DBManagerLocation(DBManager):
    def __init__(self, db):
        super().__init__(db)

    #  Insert value for location, can throw IntegrityError if item already exists for a UNIQUE column
    def insert_value(self, value):
        try:
            self.conn.execute(f'''INSERT INTO {self.table} (location) VALUES (?)''', (value,))
            self.conn.commit()
        except sqlite3.IntegrityError:
            # Since value already exists do nothing
            print("Value already exists!")

    def insert_values(self, values):
        values = [(value,) for value in values]
        self.conn.executemany(f'''INSERT INTO {self.table} (location) VALUES (?)''', values)
        self.conn.commit()

    def get_locations(self):
        return [location_tuple[0] for location_tuple in self.get_all()]
