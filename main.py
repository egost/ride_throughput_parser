import os
import re

import dateutil
import pandas as pd
from openpyxl import load_workbook


ROWS = []




def valid_times(cells):
    """Get the valid hours for the excel spreadsheet

    The columns are not always the same on different files. This is why finding the headers for the times was a good solution to find the correct columns were data should be extracted.
    """

    hours = []
    for row in cells.rows:
        for cell in row:
            if cell.row == 12: # 12 is hardcoded because it is the row where the time headers are found
                if cell.value is not None:
                    # print(cell.value)
                    hours.append(cell)

    return hours


def throughput(cells, times):
    """Get the throughput for all columns"""

    ride_throughput = {}
    for time in times:
        value_coordinate = str(time.column) + '20' # 20 is hardcoded to the row number that corresponds to the Ride Throughput
        value = cells[value_coordinate].value
        if value != 'Ride Throughput':
            ride_throughput[time.value] = value

    return ride_throughput


def get_files(directory, inc_ext=['xlsx']):
    """Gets file names of file types from a directory"""

    file_names = [
            filename \
            for filename in os.listdir(os.path.realpath(directory)) \
            if any(filename.endswith(ext) for ext in inc_ext)
            ]

    return file_names


def sweep_cells(cells):
    times = valid_times(cells)
    pkt = throughput(cells, times)

    return pkt


def sweep_sheets(workbook, date):
    """Looks through the sheets of a given document"""

    sheet_names = workbook.sheetnames

    for sheet_name in sheet_names:
        cells = workbook[sheet_name]
        ride_name = cells['A6'].value # A6 is hardcoded to the ride_name on the sheet
        pkt = sweep_cells(cells)

        row = {'ride_name': ride_name, 'date': date }

        #TODO: replace this with more pythonic solution
        #TODO: figure out why row.update({}) was rendering None
        for time, value in pkt.items():
            row[time]=value

        ROWS.append(row)


def sweep_documents(directory):
    """Looks through the documents of a given document"""
    file_names = get_files(directory)
    for filename in file_names:
        print()
        print('--------------------------------------------')
        print('Loading ' + filename)
        print('--------------------------------------------')
        wb = load_workbook(filename = os.path.join(directory, filename))

        # extract date from filename
        raw_date = re.sub('[DW_.xlsx]', '', filename)
        date = dateutil.parser.parse(raw_date)
        sweep_sheets(wb, date)


def save(df, filename):
    """Write DataFrame to a file"""
    df.to_excel(filename)


def main():
    """Sweeps through documents to extract information regarding the throughput of rides"""

    sweep_documents('attraction-operational-readiness-reports')
    # sweep_documents('test')

    # ROWS is modified by sweep_documents before this step
    df = pd.DataFrame(ROWS)

    TITLES = ['ride_name',
              'date',
              '9a-10a',
              '10a-11a',
              '11a-12p',
              '12p-1p',
              '1p-2p',
              '2p-3p',
              '3p-4p',
              '4p-5p',
              '5p-6p',
              '6p-7p',
              '7p-8p',
              '8p-9p',
              '9p-10p',
              '10p-11p'
              ]

    df = df[TITLES]
    sorted_df = df.sort_values(by=['ride_name','date'])
    save(sorted_df, 'testing.xlsx')


if __name__ == '__main__':
    main()
