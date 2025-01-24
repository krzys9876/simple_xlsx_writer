import getpass
import os

from simple_xlsx_writer import writer
from simple_xlsx_writer import postgresql_handler

def main():
    username = input("username: ")
    password = getpass.getpass()
    host = input("host: ")
    port = int(input("port: "))
    dbname = input("database name: ")

    # verify connection
    print("db time: "+postgresql_handler.get_sysdate(username,password,host,port,dbname).strftime("%Y-%m-%d %H:%M:%S"))

    base_path = os.path.dirname(__file__)

    writer.write_dummy(base_path, "dummy02")

    # fetch all tables' metadata
    query = "select * from information_schema.tables"
    postgresql_handler.write_query(query, base_path, "all_tables_pg", username, password, host, port, dbname)


if __name__ == '__main__':
    main()
