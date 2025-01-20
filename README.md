# Yet another Python XLSX writer

... yes, this is reinventing the wheel again and again, but ...

So I decided to write my own *xlsx* export for two reasons:

First and foremost, the two existing engines I use (*openpyxl*, *xlsxwriter*) available in *pandas* do not store large files efficiently.
The problem is when I must load large number of records, sometimes beyond Excel limit (2^20 = 1048576) and then send it over email 
(this is quite often the easies way to share data...). The files get way too big.

Secondly, I just want to understand *xlsx* internals and use the simples possible code to handle files. 
As a side effect, it is simpler and faster than using some other libraries.

As a simple benchmark consider a sample file of 700+k records and 18 columns. 
Standard *pandas* creates files of about **40**MB. The simple_xls_writer's file is as small as **8**MB which makes it more *email friendly*. 


(Of course when saving modified file it gets much bigger but this not the point). 

## Usage

The project consists of submodules: 

- *writer*
- *oracle_handler*. 

### writer

This generic module exposes function(s) to write raw data 
(array of arrays) into Excel file. 

This should be clear when reading the helper function:

    def write_dummy(base_path: str, target_name: str) -> None:
        data = [["A", "B", "C"], ["TEST", 1.23, "2024-10-01 12:34:56"], ["TEST", 200, "2024-10-01 12:34:56"]]
        write_raw_data(base_path, target_name, data)

Note that the only supported data types are: *str*, *int* and *float*, which relates to the way data is saved in xlsx file.

So you may have to prepare the input array yourself or use other submodules (see below).

There's a helper function *write_dummy* that saves predefined tiny file under given name.

### oracle_handler 

If you use Oracle database you can use helper method that reads query result into required structure.

First of all you may wish to verify connection. I prefer to do it this way:

    print("db time: "+oracle_handler.get_sysdate(username,password,dh_url).strftime("%Y-%m-%d %H:%M:%S"))

To save query results simply run:

    oracle_handler.write_oracle_query(query,base_path, "all_tables",username,password,dh_url)

#### Example

See: *main.py*

    ...    

    username = input("username: ")
    password = getpass.getpass()
    dh_url = input("DSN: ")
    
    # verify connection
    print("db time: "+oracle_handler.get_sysdate(username,password,dh_url).strftime("%Y-%m-%d %H:%M:%S"))

    # fetch all tables' metadata
    query = "select * from all_tables"
    base_path = os.path.dirname(__file__)
    oracle_handler.write_oracle_query(query,base_path, "all_tables",username,password,dh_url)

    ...

### CSV conversion

From time to time I come across large CSV files. When you want to load it to Excel using Power Query it generally works fine,
but what if you have to load it over the slow network. The solution is to load all files to relativelly small *xlsx* file.

#### Example

    ...

    custom_params = {
        "debug_info_every_rows": 100000,
        "row_limit": 1000000,
        "row_limit_exceed_strategy": "sheets" # truncate / files / sheets
    }
    base_path = os.path.dirname(__file__)
    writer.convert_csv("c:\\some_path\\A_VERY_LARGE_CSV.csv", base_path,
                       "converted_from_csv.xlsx", debug=True, custom_params=custom_params)
    
    ...

The *debug* option is particularly useful when you want to ensure that data load progresses.


### Exceeding Excel row limit

If your input data exceeds Excel row limit or you just want to divide it into smaller chunks, 
you may use two strategies (see: *configuration*):
    
- use files - save each chunk into a separate *xlsx* file
- use sheets - save each chunk into a separate worksheet within a single *xlsx* file

# Configuration

You may customize operations using *custom_params* parameter. The available options are:

| **Parameter**                        | **Default value**  | **Description**                                                       |
|--------------------------------------|--------------------|-----------------------------------------------------------------------|
| sheet_name                           | data               | sheet name or sheet name prefix for multiple sheets (numbered from 1) |
| python_date_format                   | %Y-%m-%d           | date format when converting data loaded from database                 | 
| python_datetime_format               | %Y-%m-%d %H:%M:%S  | as above but for datatime format                                      |
| python_datetime_remove_zeros         | True               | should empty time be removed (to save space and improve readability)  |
| python_datetime_remove_zeros_pattern | " 00:00:00"        | the pattern of empty time to be removed                               |
| headers                              | True               | does input data contain header row                                    |
| row_limit                            | 1048576-1 (2^20-1) | row limit (unchecked!)                                                |
| row_limit_exceed_strategy            | truncate           | what to do when dealing with data exceeding row limit                 |
|                                      |                    | - truncate - truncate data beyond row limit                           |
|                                      |                    | - files - generate multiple files (numbered from 1)                   |
|                                      |                    | - sheets - generate multiple sheets (numbered from 1)                 |
| debug_info_every_rows                | 10000              | print debug info every X rows, applies only when debug=True           |
| csv_delimiter                        | , (comma)          | CSV file delimiter                                                    |
| csv_quote                            | \" (double quote)  | CSV quote character                                                   |
| csv_encoding                         | utf-8              | CSV file encoding                                                     |

You provide custom options as Python *dict*:

    custom_params = {
        "sheet_name": "events_",
        "debug_info_every_rows": 100000,
        "row_limit": 1000000,
        "row_limit_exceed_strategy": "sheets"
    }


## Installation

Install package using pip:

    pip install simple-xlsx-writer

If you wish to use *Oracle* connectivity, add option:

    pip install simple-xlsx-writer[oracle]

To verify installation run:

    import os
    from simple_xlsx_writer import writer

    base_path = os.path.dirname(__file__) # or provide explicit path in interactive mode
    writer.write_dummy(base_path, "dummy01")

You should find *dummy01.xlsx* file in a given containig:

| A    | B    | C                   |
|------|------|---------------------|
| TEST | 1,23 | 2024-10-01 12:34:56 |
| TEST | 200  | 2024-10-01 12:34:56 |

## Memory considerations

Keep in mind that the module is performance and size optimized, NOT memory optimized.
The complete dataset (either from database of CSV file) is loaded to memory in order to calculate shared strings. 

For extremely large files (as for Excel terms, not extermely in general) the memory requirements are substantial.

One real-life example is conversion of 1GB CSV file (2.5M rows and 60 columns with lots of text repetitions, Windows 11). 

Loading the file to memory takes **~6GB**, slicing it to separate sheets bumps memory for a brief moment up to even **10GB**.
The resulting single xlsx file is about **60MB** (6% of original). The process takes about **8 minutes** (i5 13gen).
