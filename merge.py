#!/usr/bin/env python3

import typing
import io
import argparse
import xlrd
import xlsxwriter
import mmap
import functools


def validate_file_headers(*books: typing.List[xlrd.Book], unique_on=None):
    first_book, *rest_books = books
    validated = True

    first_sheet = first_book.sheet_by_index(0)
    known_header = [r.value for r in first_sheet.row(0)]

    if unique_on:
        print(unique_on)
        for unique_column in unique_on:
            if unique_column not in known_header:
                print(f"{unique_column} is not a known column!")

    for book in rest_books:
        sheet = book.sheet_by_index(0)
        header = [r.value for r in sheet.row(0)]

        mismatches = []
        for cell_a, cell_b in zip(sorted(known_header), sorted(header)):
            if cell_a != cell_b:
                mismatches.append((cell_a, cell_b))

        if mismatches:
            validated = False
            print(f"Mismatched headers in {book}")
            import pdb; pdb.set_trace()
            print(mismatches)


def merge_excel_files(*books, output=None, unique_on: typing.List[str] = None):
    if not output:
        output = "output.xlsx"

    get_unique_value = lambda rowdict: tuple(rowdict.values())
    if unique_on:
        get_unique_value = lambda rowdict: tuple(rowdict[col] for col in unique_on)

    hashset = set()

    output_book = xlsxwriter.Workbook(output)
    output_sheet = output_book.add_worksheet()
    output_row = 0

    book = books[0]
    sheet = book.sheet_by_index(0)
    header = [r.value for r in sheet.row(0)]

    output_sheet.write_row(output_row, 0, header)
    output_row += 1

    for book in books:
        sheet = book.sheet_by_index(0)

        rows = sheet.get_rows()
        header = [r.value for r in next(rows)]

        duplicate_rows = 0
        rows_written = 0

        for row in map(lambda row: dict(zip(header, [r.value for r in row])), rows):
            if (unique_value := get_unique_value(row)) not in hashset:
                hashset.add(unique_value)

                output_sheet.write_row(output_row, 0, row.values())
                output_row += 1
                rows_written += 1
            else:
                duplicate_rows += 1

        print(f"Wrote {rows_written}, ignored {duplicate_rows}")

    output_book.close()


open_workbook = functools.partial(xlrd.open_workbook, on_demand=True)
def open_workbook(filename):
    print(filename)
    return xlrd.open_workbook(filename, on_demand=True)

parser = argparse.ArgumentParser()
parser.add_argument(
    "--output",
    "-o",
    type=argparse.FileType("wb"),
    help="Output to store merged rows in",
)
parser.add_argument("--unique-on", type=functools.partial(str.split, sep=","))
parser.add_argument(
    "files",
    type=open_workbook,
    nargs="+",
    help="Excel files to merge (first sheet only)",
)


if __name__ == "__main__":
    args = parser.parse_args()
    validate_file_headers(*args.files)
    merge_excel_files(*args.files, output=args.output, unique_on=args.unique_on)
