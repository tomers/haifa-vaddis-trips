#!/usr/bin/env python3
import csv
import re
from pathlib import Path
import openpyxl
import phonenumbers
import validators
from bidi.algorithm import get_display

INPUT_XLSX_FILE = Path.cwd().joinpath('input.xlsx')
OUTPUT_CSV_FILE = Path(str(INPUT_XLSX_FILE) + '.csv')
XLSX_COLS = dict(name=1, phone=3, email=4, trips=5)
NAME_PREFIX = '000 haifa-vaddis'


def get_row_fields(row):
    return {field: re.sub(r'\s+', ' ', str(row[idx].value).strip())
            for field, idx in XLSX_COLS.items()}


def normalize_phone(phone, region=None):
    try:
        phone = phonenumbers.parse(phone, region)
        phone = phonenumbers.format_number(
            phone, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
    except phonenumbers.phonenumberutil.NumberParseException as e:
        if region is not None:
            return normalize_phone(phone)
        return None
    return str(phone)


def normalize_email(email):
    return email if validators.email(email) else None


def normalize_trips(trips):
    return {trip.strip() for trip in trips.split(',')}


class Person:
    def __init__(self, row):
        for field, value in get_row_fields(row).items():
            setattr(self, field, value)
        self.phone = normalize_phone(self.phone, 'IL')
        self.email = normalize_email(self.email)
        self.trips = normalize_trips(self.trips)


def convert_xlsx():
    print(f"Converting {INPUT_XLSX_FILE}")

    # read Excel file
    workbook = openpyxl.load_workbook(filename=INPUT_XLSX_FILE)
    sheet = workbook.active

    persons = [Person(sheet[row]) for row in range(2, sheet.max_row + 1)]
    trips = set.union(*(person.trips for person in persons))
    print(f'There are {len(trips)} trips:')
    for trip in trips:
        print(get_display(trip))

    # store CSV file
    with open(OUTPUT_CSV_FILE, 'w', newline='') as f:
        c = csv.writer(f)
        c.writerow(['First Name', 'Mobile Phone', 'E-mail Address'])
        for trip in trips:
            for person in persons:
                if trip not in person.trips:
                    continue
                name = f'{NAME_PREFIX} {trip} {person.name}'
                c.writerow([name, person.phone, person.email])


def main():
    convert_xlsx()


if __name__ == '__main__':
    main()
