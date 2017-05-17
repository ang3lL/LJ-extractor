#!/usr/bin/python2

import argparse
import fitz
import sys
from olevba import *
import re


def usage():
    parser = argparse.ArgumentParser(description='Extract links from Locky, Jaff PDF')
    parser.add_argument('file', action='store', help="PDF file")
    parser.add_argument("-v", "--verbose", help="Increase verbosity -v, -vv", action="count")

    return parser.parse_args()


def extract_docm(filename):

    doc = fitz.open(filename)
    count = doc.embeddedFileCount

    if count != 1:
        print "[!] ERROR: None or more than 1 embedded file."
        sys.exit(1)

    # get embedded file name
    name = doc.embeddedFileInfo(0)["name"]

    if args.verbose > 1:
        print "[+] Embedded {:} found".format(name)

    # extract embedded file
    buf = doc.embeddedFileGet(name)
    out_path = "{:}-{:}".format(args.file, name)

    # write doc
    with open(out_path, "wb") as out:
        out.write(buf)

    if args.verbose > 1:
        print "[+] Embedded file extracted: {}".format(out_path)

    return out_path


def extract_macros(filename):
    SEP = 'V'
    REGEX = {
            "links_replace": "(?:Replace)\(\"(\S+{:}\S+)\",\s\"(\S+?)\",\s\"(\S+?)\"".format(SEP),
            "links_split": "(?:Split)\(\"(\S+{:}\S+)\"".format(SEP)
            }

    # Retrive docm macros
    vba_parser = VBA_Parser_CLI(filename)
    vba_parser.extract_all_macros()

    if args.verbose > 1:
        print "[+] VBA Macros extracted"

    # search links in all vba code
    for (subfilename, stream_path, vba_filename, vba_code) in vba_parser.modules:

        matches = grep(vba_code, REGEX["links_replace"])
        # match replace
        if matches:
            links = matches[0][0].replace(matches[0][1], matches[0][2])
            links = links.split(SEP)
            continue

        matches = grep(vba_code, REGEX["links_split"])
        # match split
        if matches:
            links = matches[0].split(SEP)
            continue

    # print links
    if args.verbose == 1:
        for i in links:
            print i
    elif args.verbose > 1:
        if links:
            print "[+] Links found: {:}".format(links)
        else:
            print "[-] No links found: {}".format(links)


def grep(data, regex):

    pattern = re.compile(regex)

    try:
        return pattern.findall(data)
    except:
        return None

def main():
    global args
    args = usage()

    # extract embedded file
    doc = extract_docm(args.file)

    # extract macros from doc
    extract_macros(doc)


if __name__ == '__main__':
    main()
