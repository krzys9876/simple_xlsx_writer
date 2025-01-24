import getpass
import os

from simple_xlsx_writer import writer
from simple_xlsx_writer import oracle_handler

def main():
    username = input("username: ")
    password = getpass.getpass()
    dh_url = input("DSN: ")

    print("db time: "+oracle_handler.get_sysdate(username,password,dh_url).strftime("%Y-%m-%d %H:%M:%S"))

    base_path = os.path.dirname(__file__)

    writer.write_dummy(base_path, "dummy01")

    query = "select * from all_tables"
    oracle_handler.write_query(query, base_path, "all_tables_ora", username, password, dh_url)


if __name__ == '__main__':
    main()
