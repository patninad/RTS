import sqlite3 as sq

class DBManager:

    def __init__(self, db):
        self.set_db(db)
        self.conn = self.__setup()

        # Set the first table as default if exists
        self.tables = self.get_tables()
        if self.tables:
            self.set_table(self.tables[0])

    def __setup(self):
        conn = sq.connect(self.db)
        return conn

    def get_dbs(self):
        print(self.conn.execute('''PRAGMA database_list''').fetchall())

    def get_tables(self):
        table_tuples = self.conn.execute('''SELECT name FROM sqlite_master WHERE type = "table" AND name NOT LIKE "sqlite_%"''').fetchall()
        tables = [table_tuple[0] for table_tuple in table_tuples]
        return tables

    def get_all(self):
        return self.conn.execute(f'''SELECT * FROM {self.table}''').fetchall()
    
    def create_table(self, table_name, props):
        try:
            self.conn.execute(f'''CREATE TABLE {table_name} ({props})''')
            self.set_table(table_name)
            self.conn.commit()
        except Exception as e:
            print(f"Exception: {e}")

        return table_name

    def insert_value(self):
        pass

    def describe_table(self):
        print(self.conn.execute(f"SELECT sql FROM sqlite_master WHERE name = ?", (self.table,)).fetchall()[0][0])

    # Close connection
    def close_conn(self):
        self.conn.close()
    
    # SETTERS --------------------------------------------------
    def set_table(self, table):
        self.table = table
    
    def set_db(self, db):
        self.db = db
