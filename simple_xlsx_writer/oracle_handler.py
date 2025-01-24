import sys

import oracledb
import datetime
from simple_xlsx_writer import common


# See:
# https://cjones-oracle.medium.com/using-python-oracledb-1-0-with-sqlalchemy-pandas-django-and-flask-5d84e910cb19
def __init_oracle_version__():
    oracledb.version = "8.3.0"
    sys.modules["cx_Oracle"] = oracledb


# a helper function to verify connection
def get_sysdate(user: str, password: str, dsn: str) -> datetime.datetime:
    __init_oracle_version__()
    with oracledb.connect(user=user, password=password, dsn=dsn) as connection:
        with connection.cursor() as cursor:
            res = cursor.execute("select sysdate from dual").fetchone()
            return res[0]


def get_data_from_query(query: str, user: str, password: str, dsn: str, debug: bool = False, custom_params = None) -> []:
    __init_oracle_version__()

    with oracledb.connect(user=user, password=password, dsn=dsn) as connection:
        data = common.get_data_from_query(connection, query, debug, custom_params)
    return data


def write_query(query: str, base_path: str, target_file_name: str, user: str, password: str, dsn: str,
                debug: bool = False, custom_params = None) -> None:
    data = get_data_from_query(query,user,password,dsn, debug, custom_params)
    common.write_data(base_path, target_file_name, data, debug, custom_params)
