import psycopg
import datetime
from simple_xlsx_writer import common
import urllib.parse



def __connection_string__(user: str, password: str, host: str, port: int, dbname: str) -> str:
    # postgresql://[userspec@][hostspec][/dbname][?paramspec]
    return f"postgresql://{user}:{urllib.parse.quote_plus(password)}@{host}:{str(port)}/{dbname}"


# a helper function to verify connection
def get_sysdate(user: str, password: str, host: str, port: int, dbname: str) -> datetime.datetime:
    with psycopg.connect(__connection_string__(user, password, host, port, dbname)) as connection:
        return common.get_first_record(connection, "select now()")


def get_data_from_query(query: str, user: str, password: str, host: str, port: int, dbname: str, debug: bool = False,
                        custom_params: dict = None) -> []:
    with psycopg.connect(__connection_string__(user, password, host, port, dbname)) as connection:
        data = common.get_data_from_query(connection, query, debug, custom_params)
    return data


def write_query(query: str, base_path: str, target_file_name: str, user: str, password: str, host: str, port: int, dbname: str,
                debug: bool = False, custom_params = None) -> None:
    data = get_data_from_query(query,user,password,host, port, dbname, debug, custom_params)
    common.write_data(base_path, target_file_name, data, debug, custom_params)
