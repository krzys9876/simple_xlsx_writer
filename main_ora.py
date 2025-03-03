import getpass
import os

from simple_xlsx_writer import writer
from simple_xlsx_writer import oracle_handler

def main():
    username = input("username: ")
    password = getpass.getpass()
    dsn = input("DSN: ")
    mode = input("mode (thick/thin): ")

    custom_params = {'oracle_client_mode': mode}

    # verify connection
    print("db time: "+oracle_handler.get_sysdate(username, password, dsn, custom_params).strftime("%Y-%m-%d %H:%M:%S"))

    base_path = os.path.dirname(__file__)

    writer.write_dummy(base_path, "dummy01")

    # fetch all tables' metadata
    query = "select * from all_tables"
    oracle_handler.write_query(query, base_path, "all_tables_ora", username, password, dsn,
                               debug = True, custom_params = custom_params)


if __name__ == '__main__':
    main()
