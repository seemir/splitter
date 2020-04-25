# -*- coding: utf-8 -*-

__author__ = 'Samir Adrik and Mohamed Adrik'
__email__ = 'samir.adrik@gmail.com, mohamed.adrik@knowit.no'
__version__ = "0.1.4"

import os
import sys
import datetime
from time import time

from pyfiglet import Figlet
from argparse import ArgumentParser
from pandas import read_csv, DataFrame

from win32com.client import Dispatch


def splitter(args):
    """
    Naive / brute-force method that splits large .csv file into n separate .xlsx files based on a
    row breakpoint (n), i.e. each row is appended to its own file until row_num % n == 0. Then the
    algorithm starts over and appends the next series of rows into the same created files starting
    from file num 1, 2, 3 ... n, as follows:

    csv_row[1]    -> append -> xlsx_file[1]
    csv_row[2]    -> append -> xlsx_file[2]
    ...
    csv_row[n]    -> append -> xlsx_file[n]

    csv_row[n+1]  -> append -> xlsx_file[1]
    csv_row[n+2]  -> append -> xlsx_file[2]
    ...
    csv_row[2n]   -> append -> xlsx_file[n]

    csv_row[2n+1] -> append -> xlsx_file[1]
    csv_row[2n+2] -> append -> xlsx_file[2]
    ...
    csv_row[3n]   -> append -> xlsx_file[n]
    .
    .
    .
    csv_row[n^3]   -> append -> xlsx_file[n]

    Parameters
    ----------
    args        : arguments (see below)
                  csv, str - name of file / file path
                  n, int - number of files / row break
                  na, str, optional - string to replace NaN values, default is NULL

    Notes
    -----
    If other that .csv file is passed to method a TypeError is raised

    """
    try:
        file_name = args.csv
        n = args.n
        na = args.na if args.na else "NULL"

        start = time()
        full_df = read_csv(file_name, delimiter=";").fillna(value=na).astype(str)
        headers = list(full_df.head(0))
        df_list = [DataFrame(columns=headers) for _ in range(n)]
        row_num = len(full_df.index)

        count = 0
        progress = 1
        bar = []

        for row in full_df.iterrows():
            percent = (progress / row_num) * 100
            if percent % 5 == 0:
                bar.append("#")

            sys.stdout.write("\rprocessing '{}' [{} row nr. {} ({}%)]".format(
                file_name, "".join(bar), progress, round(percent, 3)))
            sys.stdout.flush()

            if count == n:
                count = 0
            df_list[count] = df_list[count].append(row[1])

            count += 1
            progress += 1

        for df in df_list:
            for header in headers:
                if header == "MeteringPointId":
                    continue
                df[header] = "'" + df[header]

        print("\n\nfinished processing '{}' file, elapsed: {}s\n".format(
            file_name, round((time() - start), 7)))

        timestamp = datetime.datetime.now().isoformat().replace(".", "_").replace(":", "_")
        file_dir = os.path.dirname(os.path.abspath(__file__)) + r"\ETA_{}".format(timestamp)

        if not os.path.exists(file_dir):
            os.makedirs(file_dir)

        excel = Dispatch("Excel.Application")
        for i, df in enumerate(df_list):
            saved_file_name = r"{}\ETA_{:02d}.xlsx".format(file_dir, i + 1)
            sys.stdout.write("\rsaving '{}' (#{})".format(saved_file_name, i + 1))
            sys.stdout.flush()

            df.reset_index(drop=True).to_excel(saved_file_name, index=False,
                                               sheet_name="ETA_{:02d}".format(i + 1))
            wb = excel.Workbooks.Open(saved_file_name)
            excel.Worksheets(1).Activate()
            excel.ActiveSheet.Columns.AutoFit()
            excel.ActiveSheet.Columns.HorizontalAlignment = -4131
            excel.ActiveSheet.Columns.Replace(".0", "")
            excel.ActiveSheet.Columns.Replace("'", "")
            wb.Save()
            wb.Close()

        print("\n\nfinished creating and saving n = {} '.xlsx' file(s), elapsed: {}s".format(
            n, round((time() - start), 7)))
    except Exception as reading_exception:
        raise OSError(
            "Something happened! please insure that the input file-format is '.csv', "
            "exited with'{}'".format(reading_exception))


def main():
    """
    main program, i.e. primary entrance point to the application

    """
    fig = Figlet(font='standard')
    print(fig.renderText('splitter'))
    print('Authors: ' + __author__)
    print('Email: ' + __email__)
    print('Version: ' + __version__ + '\n')

    parser = ArgumentParser(description="splits large .csv file into n separate .xlsx files "
                                        "based on row breakpoint (n)")
    parser.add_argument("-csv", help="name or path to csv file to process", dest="csv", type=str,
                        required=True)
    parser.add_argument("-n", help="number of .xlsx files to produce, also the row break point",
                        dest="n", type=int, required=True)
    parser.add_argument("-na",
                        help="optional, string representation for NaN values, default is 'NULL'",
                        dest="na", type=str, required=False)
    parser.set_defaults(func=splitter)
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
