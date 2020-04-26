# Splitter

Splitter is a naive / brute-force python library for dealing with splitting `.csv` files containing data 
into `n` separate `.xlsx` files. 

# Description 
The library contains a `splitter()` function which takes in three arguments, i.e. `csv`, `n` and `na`[optional]. 
Given these three arguments the function splits the full `.csv` file with data into `n` separate `.xlsx` files 
based on a row breakpoint (which is `n`), i.e. each row is appended to its own file until `row_num % n == 0`. Then the 
algorithm starts over and appends the next series of rows into the already created files starting over from file number 
`1, 2, 3 ... n`, as follows:
```
csv_row[1]     -> append -> xlsx_file[1]
csv_row[2]     -> append -> xlsx_file[2]
...
csv_row[n]     -> append -> xlsx_file[n]

csv_row[n+1]   -> append -> xlsx_file[1]
csv_row[n+2]   -> append -> xlsx_file[2]
...
csv_row[2n]    -> append -> xlsx_file[n]

csv_row[2n+1]  -> append -> xlsx_file[1]
csv_row[2n+2]  -> append -> xlsx_file[2]
...
csv_row[3n]    -> append -> xlsx_file[n]
.
.
.

```


## Installation

Install dependencies with `--user` permission

```bash
git clone repo
cd to the/splitter/folder
pip install -r requirements.txt --user

```

## Usage
As the library is not published to PyPi one cannot install the library using the standard python package manager 
[pip](https://pip.pypa.io/en/stable/). Still, assuming one has the source code the following command can be used 
to run the program.

```bash
python splitter.py -csv csv_file_to_process.csv -n number_of_xlsx_files

```
## Support

For more information about the program and `args`, use the flag `-h` or `--help`:

```bash
python splitter.py -h

           _ _ _   _
 ___ _ __ | (_) |_| |_ ___ _ __
/ __| '_ \| | | __| __/ _ \ '__|
\__ \ |_) | | | |_| ||  __/ |
|___/ .__/|_|_|\__|\__\___|_|
    |_|

Authors: Samir Adrik and Mohamed Adrik
Email: samir.adrik@gmail.com, mohamed.adrik@knowit.no
Version: 0.1.5

usage: splitter.py [-h] -csv CSV -n N [-na NA]

splits large .csv file into n separate .xlsx files based on row breakpoint (n)

optional arguments:
  -h, --help  show this help message and exit
  -csv CSV    name or path to csv file to process
  -n N        number of .xlsx files to produce, also the row break point
  -na NA      optional, string representation for NaN values, default is
              'NULL'

```

## License
[MIT](https://choosealicense.com/licenses/mit/)
