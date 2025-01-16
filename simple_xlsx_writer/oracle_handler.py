from types import NoneType
import sys

import oracledb
import datetime
from simple_xlsx_writer import writer


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


def get_data_from_query(query: str, user: str, password: str, dsn: str, headers: bool = True) -> []:
    __init_oracle_version__()

    data = []
    with oracledb.connect(user=user, password=password, dsn=dsn) as connection:
        with connection.cursor() as cursor:
            result = cursor.execute(query)

            if headers:
                row = []
                for c in result.description:
                    row.append(c[0])
                data.append(row)

            for r in result:
                row=[]
                for cell in r:
                    if type(cell)==int or type(cell)==float or type(cell)==str:
                        row.append(cell)
                    elif type(cell)==datetime.datetime:
                        row.append(cell.strftime("%Y-%m-%d %H:%M:%S").replace(" 00:00:00", ""))
                    elif type(cell) == datetime.date:
                        row.append(cell.strftime("%Y-%m-%d"))
                    elif type(cell)==NoneType:
                        row.append("")
                    else:
                        raise TypeError(f"Unsupported data type found in cell {cell} of type {type(cell)}")
                data.append(row)
    return data


def write_oracle_query(query: str, base_path: str, target_file_name: str, user: str, password: str, dsn: str,
                       header: bool = True, debug: bool = False) -> None:
    if debug:
        print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ ": executing query")
    data = get_data_from_query(query,user,password,dsn,header)
    if debug:
        print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ ": writing file")
    writer.write_raw_data(base_path, target_file_name, data, debug)
    if debug:
        print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ ": finished")

