# Yet another Python XLSX writer

... yes, this is reinventing the wheel again and again, but ...

So I decided to write my own *xlsx* export for two reasons:

First and foremost, the two existing engines I use (*openpyxl*, *xlsxwriter*) available in *pandas* do not store large files efficiently.
The problem is when I must load large number of records, up to Excel limit (2^20 = 1048576) and then send it over email 
(this is quite often the easies way to share data...). The files get way big.

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

So you may have to prepare the input array yourself or use other submdules (see below).

There's a helper function *write_dummy* that saves predefined tiny file under given name.

### oracle_handler 

If you use Oracle database you can use helper method that reads query result into required structure.

First of all you may wish to verify connection. I prefer to do it this way:

    print("db time: "+oracle_handler.get_sysdate(username,password,dh_url).strftime("%Y-%m-%d %H:%M:%S"))

To save query results simply run:

    oracle_handler.write_oracle_query(query,base_path, "all_tables",username,password,dh_url)

#### Example

See: *main.py*

    username = input("username: ")
    password = getpass.getpass()
    dh_url = input("DSN: ")
    
    # verify connection
    print("db time: "+oracle_handler.get_sysdate(username,password,dh_url).strftime("%Y-%m-%d %H:%M:%S"))

    # fetch all tables' metadata
    query = "select * from all_tables"
    base_path = os.path.dirname(__file__)
    oracle_handler.write_oracle_query(query,base_path, "all_tables",username,password,dh_url)
