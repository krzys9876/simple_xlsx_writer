import getpass
import os

from simple_xlsx_writer import writer
from simple_xlsx_writer import oracle_handler

def main():
    username = input("username: ")
    password = getpass.getpass()
    dsn = input("DSN: ")

    # verify connection
    print("db time: "+oracle_handler.get_sysdate(username,password,dsn).strftime("%Y-%m-%d %H:%M:%S"))

    base_path = os.path.dirname(__file__)

    writer.write_dummy(base_path, "dummy01")

    # fetch all tables' metadata
    query = "select * from all_tables"
    oracle_handler.write_query(query, base_path, "all_tables_ora", username, password, dsn)


if __name__ == '__main__':
    main()
