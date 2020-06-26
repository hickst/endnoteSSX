#!/usr/bin/env python3
#
# Module to extract endnote data to a CSV spreadsheet.
#   Written by: Tom Hicks. 6/25/20.
#   Last Modified: Initial creation.
#
import os
import sys
import argparse
import csv
import xml.etree.ElementTree as xmlet


# Separator character for fields with multiple values
FIELD_SEPARATOR = ';'

# Program name for this tool.
PROG_NAME = 'ssx'

# Version of this tool.
VERSION = '0.1.0'


def main (argv=None):
    """
    The main method for the tool. This method is called from the command line,
    processes the command line arguments and calls into the ImdTk library to do its work.
    This main method takes no arguments so it can be called by setuptools.
    """

    # the main method takes no arguments so it can be called by setuptools
    if (argv is None):                      # if called by setuptools
        argv = sys.argv[1:]                 # then fetch the arguments from the system

    # setup command line argument parsing and add shared arguments
    parser = argparse.ArgumentParser(
        prog=PROG_NAME,
        formatter_class=argparse.RawTextHelpFormatter,
        description='Extract endnote data to a CSV spreadsheet.'
    )

    parser.add_argument(
        '-in', '--input_file', dest='input_file', required=True, metavar='filepath',
        help='Path to a readable input Endnote XML dump.'
    )

    parser.add_argument(
        '-out', '--output_file', dest='output_file',required=True,  metavar='filepath',
        help='File path of file to hold the CSV output file.'
    )

    # actually parse the arguments from the command line
    args = vars(parser.parse_args(argv))

    # if debugging, set verbose and echo input arguments
    if (args.get('debug')):
        args['verbose'] = True              # if debug turn on verbose too
        print("({}.main): ARGS={}".format(PROG_NAME, args), file=sys.stderr)

    # if input file path given, check the file path for validity
    input_file = args.get('input_file')
    check_input_file(input_file, PROG_NAME) # may system exit here and not return!
    output_file = args.get('output_file')

    # add additional arguments to args
    args['PROG_NAME'] = PROG_NAME
    args['VERSION'] = VERSION

    # TODO: do the work
    if (args.get('verbose')):
        print("{} called with arguments: {}".format(PROG_NAME, args))

    root = read_xml(input_file)
    lines = []
    for rec in root.findall('./records/record'):
        line = {}

        auths = [ auth.text for auth in rec.findall('./contributors/authors/author/style') ]
        line['Authors'] = ('; ').join(auths)

        pubDate = rec.find('./dates/year/style')
        line['PubDate'] = pubDate.text if ((pubDate is not None) and (pubDate.text is not None)) else ''
    
        journal = rec.find('./titles/secondary-title/style')
        line['Journal'] = journal.text if ((journal is not None) and (journal.text is not None)) else ''

        title = rec.find('./titles/title/style')
        line['Title'] = title.text if ((title is not None) and (title.text is not None)) else ''

        vol = rec.find('./volume/style')
        volume = vol.text if ((vol is not None) and (vol.text is not None)) else ''
        num = rec.find('./number/style')
        number = num.text if ((num is not None) and (num.text is not None)) else ''
        pgs = rec.find('./pages/style')
        pages = pgs.text if ((pgs is not None) and (pgs.text is not None)) else ''
        vol_ref = "{}:{}:{}".format(volume, number, pages)
        line['Volume'] = vol_ref if (vol_ref != '::') else ''

        label = rec.find('./label/style')
        line['Label'] = label.text if ((label is not None) and (label.text is not None)) else ''

        work_type = rec.find('./work-type/style')
        line['WorkType'] = work_type.text if ((work_type is not None) and (work_type.text is not None)) else ''

        # save the current line in the lines array
        lines.append(line)

    save_to_CSV(lines, output_file)


def check_input_file (input_file, tool_name, exit_code=20):
    """
    If an input file path is given, check that it is a good path. If not, then exit
    the entire program here with the specified (or default) system exit code.
    """
    if (not readable_file_path(input_file)):
        errMsg = "({}): ERROR: a readable, valid Endnote XML dump file must be specified. Exiting...".format(tool_name)
        print(errMsg, file=sys.stderr)
        sys.exit(exit_code)


def readable_file_path (apath):
    """ Tell whether the given path points to a readable file or not. Follows symbolic links. """
    return (apath and os.path.isfile(apath) and os.access(apath, os.R_OK))


def read_xml (infile):
    tree = xmlet.parse(infile)
    return tree.getroot()


def save_to_CSV (lines, filename): 
    # specifying the fields for csv file 
    fields = ['Authors', 'PubDate', 'Title', 'Journal', 'Volume', 'Label', 'WorkType']
  
    # writing to csv file 
    with open(filename, 'w') as csvfile: 
        writer = csv.DictWriter(csvfile, fieldnames=fields) 
        writer.writeheader() 
        writer.writerows(lines) 



if __name__ == "__main__":
    main()
