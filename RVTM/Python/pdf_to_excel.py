#!/usr/bin/env python3
"""
pdf_to_excel.py

Extract all tables from a multi‐page PDF and consolidate them into a single clean Excel sheet.
"""

import pdfplumber
import pandas as pd
import argparse
import os

def extract_tables_from_pdf(pdf_path):
    """
    Extracts and concatenates tables from all pages of the PDF.
    Returns a pandas.DataFrame.
    """
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                tables.append(table)

    if not tables:
        raise ValueError("No tables found in PDF.")

    # Use the first row of the first table as header
    header = tables[0][0]
    rows = []

    for table in tables:
        for row in table[1:]:
            # skip duplicated header rows
            if row == header:
                continue
            rows.append(row)

    # Build DataFrame
    df = pd.DataFrame(rows, columns=header)
    return df

def clean_dataframe(df):
    """
    Optional: apply any row/column cleanup here.
    e.g. strip whitespace, replace newline chars, etc.
    """
    # Example: strip leading/trailing whitespace and replace internal newlines
    df = df.applymap(lambda cell: cell.strip().replace('\n', ' ') if isinstance(cell, str) else cell)
    return df

def main():
    parser = argparse.ArgumentParser(
        description="Convert all tables in a PDF to a single Excel sheet."
    )
    parser.add_argument("input_pdf", help="Path to the source PDF file")
    parser.add_argument("output_xlsx", help="Path for the resulting Excel file")
    args = parser.parse_args()

    if not os.path.isfile(args.input_pdf):
        print(f"Error: PDF file not found: {args.input_pdf}")
        return

    print("Extracting tables from PDF...")
    df = extract_tables_from_pdf(args.input_pdf)

    print("Cleaning up DataFrame...")
    df = clean_dataframe(df)

    print(f"Writing to Excel: {args.output_xlsx} …")
    # Ensure .xlsx extension is used
    df.to_excel(args.output_xlsx, index=False, engine="openpyxl")
    print("Done!")

if __name__ == "__main__":
    main()
