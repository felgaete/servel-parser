#!/usr/bin/jython
# -*- coding: utf-8 -*-

import os
import sys
import csv
import string
import argparse

sys.path.append("pdfbox-app-2.0.7.jar")

import org.apache.pdfbox.pdmodel.PDDocument as PDDocument
import org.apache.pdfbox.text.PDFTextStripperByArea as PDFTextStripperByArea
import java.awt.geom.Rectangle2D.Float as r2df

MAX_ROWS_PER_PAGE = 65
START_Y = 62
TEXT_HEIGHT = 8
EVEN_OFFSET = 0
ODD_OFFSET = 0


def get_pdf_document(name):
    with open(name, "rb") as f:
        pdf = f.read()
        return PDDocument.load(pdf)


def get_province_and_area(document):
    stripper = PDFTextStripperByArea()
    stripper.addRegion("region", r2df(190, 23, 150, 6))
    stripper.addRegion("area", r2df(483, 23, 150, 6))
    stripper.addRegion("province", r2df(190, 35, 150, 6))

    page = document.getPage(0)
    stripper.extractRegions(page)

    province_and_area = {}
    for region in stripper.getRegions():
        province_and_area[region] = stripper.getTextForRegion(region).strip()

    return province_and_area if province_and_area else None


def get_row(page, offset):
    stripper = PDFTextStripperByArea()

    stripper.addRegion("name", r2df(23, START_Y + offset, 224, TEXT_HEIGHT))
    stripper.addRegion("nin", r2df(295, START_Y + offset, 45, TEXT_HEIGHT))
    stripper.addRegion("sex", r2df(340, START_Y + offset, 18, TEXT_HEIGHT))
    stripper.addRegion("address", r2df(380, START_Y + offset, 210, TEXT_HEIGHT))
    stripper.addRegion("circumscription", r2df(591, START_Y + offset, 195, TEXT_HEIGHT))
    stripper.addRegion("place", r2df(790, START_Y + offset, 26, TEXT_HEIGHT))

    stripper.extractRegions(page)

    row = {}
    for region in stripper.getRegions():
        value = stripper.getTextForRegion(region).strip()
        if value:
            row[region] = value

    return row if row else None


def get_records_from_page(document, num):
    page = document.getPage(num)
    accumulated_offset = 0

    records = []
    for r in range(0, MAX_ROWS_PER_PAGE):
        row = get_row(page, accumulated_offset)

        offset = EVEN_OFFSET if r % 2 == 0 else ODD_OFFSET
        accumulated_offset += offset + TEXT_HEIGHT

        if row:
            records.append(row)

    return records


def get_filename_from_path(path):
    path = path.lower().split(os.sep)[-1]
    path = path.split(".")
    if len(path) > 1:
        path = path[:-1]
    return "".join(path)


def safe_filename(filename):
    equivalents = {
        "á": "a",
        "é": "e",
        "í": "í",
        "ó": "o",
        "ú": "u",
        "ü": "u",
        "ñ": "ni",
        "'": ""
    }
    whitelist = "-_.() %s%s" % (string.ascii_letters, string.digits)

    filename = filename.lower()
    for char in equivalents:
        if char in filename:
            filename = filename.replace(char, equivalents[char])

    return ''.join(c for c in filename if c in whitelist)


def normalize_row(row):
    fields = ["name", "nin", "sex", "address", "circumscription", "place"]

    for key in fields:
        if key not in row:
            row[key] = ''
    return row


def results_to_cli(results):
    for row in results:
        row = normalize_row(row)

        print '%-60s %12s %3s %-80s %30s %5s' % (row["name"], row["nin"], row["sex"], row["address"],
                                                 row["circumscription"], row["place"])


def results_to_csv(filename, records, delimiter):
    with open("{}.csv".format(filename), 'ab') as csv_file:
        writer = csv.writer(csv_file, delimiter=delimiter)
        for row in records:
            row = normalize_row(row)
            writer.writerow([row["name"].encode("utf-8"), row["nin"], row["sex"], row["address"].encode("utf-8"),
                             row["circumscription"].encode("utf-8"), row["place"]])

def parse_pdf_document(file, start_page, end_page, output, delimiter):
    document = get_pdf_document(file)
    max_pages = document.getNumberOfPages()

    if not start_page or start_page < 0:
        start_page = 0
    else:
        start_page -= 1
    if not end_page or end_page > max_pages :
        end_page = max_pages
    if not delimiter:
        delimiter = ','

    province_and_area = get_province_and_area(document)

    filename = u"{}-{}-{}".format(
        province_and_area["region"], province_and_area["province"], province_and_area["area"]
    ) if all(k in province_and_area for k in ("region", "province", "area")) else None
    if filename is None:
        filename = get_filename_from_path(args.file)
    filename = safe_filename(filename)
    for i in range(start_page, end_page):
        records = get_records_from_page(document, i)

        if output == "cli":
            results_to_cli(records)
        elif output == "csv":
            results_to_csv(filename, records, delimiter)

    document.close()

def main(args):

    if not(args.file or args.dir):
        print "You must supply either a --file or --dir arg"
        return

    if args.file:
        parse_pdf_document(args.file, args.start, args.end, args.output, args.delimiter)
        return

    if args.dir:
        for _, _, files in os.walk(args.dir):
            for f in files:
                filename, fileext = os.path.splitext(f)
                if fileext == ".pdf":
                    print u"Parsing file {}".format(f)
                    parse_pdf_document(f, args.start, args.end, args.output, args.delimiter)
            return


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("--file", type=str,
                        help="path to pdf file")
    parser.add_argument("--start", type=int,
                        help="Start from page")
    parser.add_argument("--end", type=int,
                        help="Until to page")
    parser.add_argument("--output", type=str, choices=["csv", "cli"], default="cli",
                        help="Get numbers of pages")
    parser.add_argument('--delimiter', type=str,
                        help="Set csv delimiter")
    parser.add_argument('--dir', type=str, nargs='?', const='.',
                        help="Parse all files in a directory")
    args = parser.parse_args()

    main(args)
