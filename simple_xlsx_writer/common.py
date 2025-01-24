from simple_xlsx_writer import writer
import decimal
import datetime
from types import NoneType


def get_first_record(connection, query: str):
    with connection.cursor() as cursor:
        with connection.cursor() as cursor:
            res = cursor.execute(query).fetchone()
            return res[0]


def get_data_from_query(connection, query: str, debug: bool = False, custom_params: dict = None):
    params = writer.update_params(custom_params)
    data = []
    date_format = params["python_date_format"]
    datetime_format = params["python_datetime_format"]
    datetime_remove_zeros = params["python_datetime_remove_zeros"]
    datetime_remove_zeros_pattern = params["python_datetime_remove_zeros_pattern"]
    headers = params["headers"]
    with connection.cursor() as cursor:
        if debug:
            print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ": executing query")

        result = cursor.execute(query)

        if headers:
            row = []
            for c in result.description:
                row.append(c[0])
            data.append(row)

        for r in result:
            row = []
            for cell in r:
                if type(cell) == int or type(cell) == float:
                    row.append(cell)
                elif type(cell) == decimal.Decimal:
                    row.append(float(cell))
                elif type(cell) == str:
                    row.append(writer.escape_invalid_chars(cell))
                elif type(cell) == datetime.datetime:
                    txt = cell.strftime(datetime_format)
                    if datetime_remove_zeros:
                        txt = txt.replace(datetime_remove_zeros_pattern, "")
                        row.append(txt)
                elif type(cell) == datetime.date:
                    row.append(cell.strftime(date_format))
                elif type(cell) == NoneType:
                    row.append("")
                else:
                    raise TypeError(f"Unsupported data type found in cell {cell} of type {type(cell)}")
            data.append(row)

    return data


def write_data(base_path: str, target_file_name: str, data: [], debug: bool = False, custom_params: dict = None) -> None:
    if debug:
        print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ ": writing file")
    writer.write_raw_data(base_path, target_file_name, data, debug, custom_params)
    if debug:
        print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ ": finished")

