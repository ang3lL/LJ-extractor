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
    paths = []

    if count < 1:
        print "[!] ERROR: None or more than 1 embedded file."
        sys.exit(1)

    for i in range(0, count):
        # get embedded file name
        name = doc.embeddedFileInfo(i)["name"]

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

        paths.append(out_path)

    return paths


def extract_macros(filename):
    SEP = 'V'
    REGEX = {
            "links_replace": "(\S+) = (?:Replace)\(\"(\S+{:}\S+)\", \"(\S+?)\", \"(\S+?)\"".format(SEP),
            "other_replace": "(\S+) = (?:Replace)\(({:}), \"(\S+?)\", \"(\S+?)\"",
            "links_split": "(?:Split)\(\"(\S+{:}\S+)\"".format(SEP)
            }

    # Retrive docm macros
    vba_parser = VBA_Parser_CLI(filename)
    if not vba_parser.detect_vba_macros():
        return
    vba_parser.extract_all_macros()

    links = None
    # search links in all vba code
    for (subfilename, stream_path, vba_filename, vba_code) in vba_parser.modules:

        if args.verbose > 1:
            if vba_code:
                print "\t[+] VBA Macros extracted : {:}".format(vba_filename)
            else:
                print "\t[+] No VBA Macros found"

        matches = grep(vba_code, REGEX["links_replace"])
        # match replace
        if matches:
            links = matches[0][1].replace(matches[0][2], matches[0][3])
            # try to match others replace
            matches = grep(vba_code, REGEX["other_replace"].format(matches[0][0]))
            if matches:
                for match in matches:
                    links = links.replace(match[2], match[3])

            links = links.split(SEP)
            continue

        matches = grep(vba_code, REGEX["links_split"])
        # match split
        if matches:
            links = matches[0].split(SEP)
            continue

    # print links
    if args.verbose == 1:
        if links:
            for link in links:
                print link

    elif args.verbose > 1:
        if links:
            print "\t[+] Links found: {:}".format(links)
        else:
            print "\t[-] No links found"


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
    docs= extract_docm(args.file)

    for f in docs:
        if args.verbose > 1:
            print "[+] Analyzing: {:}".format(f)
        # extract macros from doc
        extract_macros(f)


if __name__ == '__main__':
    main()
