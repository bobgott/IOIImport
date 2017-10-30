import argparse
import configparser
import datetime
import logging
import ntpath
import os
import sys

import openpyxl

from applic.dicttoxml import DictToXML, DXMLGeneratedError

__project__ = 'IOI_Import'
__author__ = "Robert Gottesman"
__version_date__ = "10/29/2017"
__err_prefix__ = 'IOI'
__high_err_num__ = 13

""" Change Log
04/17/2017 - Groups column can be separated with '_' or ','
06/14/2017 - Removed DD Wiki Version text (not used)
10/29/2017 - Program starting from __init()__.py. Changed main folder name to 'files'
"""
class IOIGeneratedError(Exception):
    """
    Handle known problems in this module passing detail information
    """
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)

def valid_date(date_string):
    """ Return date/time object in YYYY-MM-DD format

    :param date_string: (str) Date in YYYY-MM-DD
    :return: date object
    """
    try:
        return datetime.datetime.strptime(date_string, "%Y-%m-%d")
    except ValueError:
        msg = "Not a valid date: '{0}'.".format(date_string)
        raise argparse.ArgumentTypeError(msg)

def print_lookup_fields(spreadsheet_info):
    """ Used for debugging (ignore)

    :param spreadsheet_info: Dictionary format of xlsx/csv file
    :return: None
    """
    top_index = spreadsheet_info['Lookups']
    for letter in sorted(top_index.keys()):
        print('{0} - Lookup Fields'.format(letter))
        # note: lookup_fields['PropertySubType'][0]['Enumeration'] .. get Enumeration
        for fld_name in top_index[letter]:
            print(fld_name)

class ResoXLSXtoDict:
    CSV_SUB_FOLDER = 'XLS-to-CSV'
    XLS_SUB_FOLDER = 'files'
    """
    This program assumes each xlsx tab is converted into a csv file with filename <sheet>.csv. Remove cr/lf
    .. Used program 'XLS to CSV Converter' to make this a one step process
    .. (http://cwestblog.com/2013/04/12/batch-excel-to-csv-converter-application/)
    """

    def __init__(self, working_folder, xlsx_filepath):
        """ Read xlsx files into internal dictionary 'spreadsheet_info' {}

        :param working_folder: (str) folder for config.ini file
        :param xlsx_filepath: (str) Full path name for input xlsx file (result xml placed in same folder)
        :return: Void. Raise IOIGeneratedError on error
        """
        self.working_folder = working_folder
        self.xlsx_filepath = xlsx_filepath
        # xml will be in same folder w/same basename but with .xml extension
        self.xml_filepath = os.path.splitext(xlsx_filepath)[0] + '.xml'
        self.xml_filename = ntpath.basename(self.xml_filepath)
        self.logger = logging.getLogger(__project__  + '.' + self.__class__.__name__)
        self.logger.debug("Initialize {0} with verson:{1}".format(self.__class__.__name__, __version_date__))
        self.resource_sheets = []
        self.lookup_sheet = None
        self.spreadsheet_info = {'Resources': {}, 'Lookups': {}}  # Container of xlsx data
        self._read_config_ini()

    def _read_config_ini(self):
        """ Read in program configuration (config.ini). See readme.md for detail

        :return: Void. Raise IOIGeneratedError on error.
        """
        # See: https://docs.python.org/3/library/configparser.html
        self.config = configparser.ConfigParser()
        self.config.optionxform = str   # Setting str, makes option names case sensitive:
        config_file = os.path.join(self.working_folder, 'config.ini')
        if not os.path.isfile(config_file):
            raise IOIGeneratedError("[IOI-06] Cannot Process Program Config INI File: " + config_file)
        try:
            self.config.read(config_file)
        except FileNotFoundError:
            raise IOIGeneratedError("[IOI-01] Cannot Process Program Config INI File: " + config_file)

        # Read in Resource Sheets
        # formatted as: [sheet tab name] = [output xml resource name]
        try:
            if 'ResourceSheets' in self.config and len(self.config['ResourceSheets']) > 0:
                resource_sheets_found = True
                for resource_sheet in self.config['ResourceSheets']:
                    self.resource_sheets.append(resource_sheet)
            else:
                resource_sheets_found = False
        except KeyError:
            raise IOIGeneratedError("[IOI-12] Cannot find 'ResourceSheets' key in config.ini")

        # Get Lookup Sheet Name (only 1 allowed)
        # formatted as: [LookupSheet] = [lookup spreadsheet tab name]
        try:
            if 'LookupSheets' in self.config and len(self.config['LookupSheets']) > 0:
                lookup_sheets_found = True
                self.lookup_sheet = self.config['LookupSheets']['LookupSheet']
            else:
                lookup_sheets_found = False
        except KeyError:
            raise IOIGeneratedError("[IOI-09] Cannot find 'LookupSheet' or child 'LookupSheets' key config.ini")

        if not resource_sheets_found and not lookup_sheets_found:
            raise IOIGeneratedError("[IOI-13] Missing sections ResourceSheets and LookupSheets in config.ini")

    def read_xlsx_file(self):
        """ Open .xlsx file and read into self.spreadsheet_info

        :return: void. Raise IOIGeneratedError on error
        """
        try:
            wb = openpyxl.load_workbook(self.xlsx_filepath)
        except FileNotFoundError:
            raise IOIGeneratedError('[IOI-07] XLSX input file {0} not found'.format(self.xlsx_filepath))
        # Read in Resource Sheet rows
        for resource_sheet in self.resource_sheets:
            try:
                ws = wb.get_sheet_by_name(name=resource_sheet)
            except KeyError:
                raise IOIGeneratedError("[IOI-11] Resource Sheet name '{0}' does not exist in .xlsx file".format(resource_sheet))
            self._create_resource_dict(resource_sheet, ws)
        # Read in Lookup Sheet rows
        if self.lookup_sheet is not None:
            try:
                ws = wb.get_sheet_by_name(name=self.lookup_sheet)
            except KeyError:
                raise IOIGeneratedError("[IOI-08] Lookup Sheet name '{0}' does not exist in .xlsx file".format(self.lookup_sheet))

            self._create_lookup_dict(ws)

            if len(self.spreadsheet_info['Lookups']) == 0:
                raise IOIGeneratedError('[W202] No Lookup Lookups Processed (tab: {})'.format(self.lookup_sheet))


    def _create_lookup_dict(self, ws):
        """ Populate xlsx lookup rows into internal dictionary (self.spreadsheet_info['Lookups'])

        :param ws: (obj) lookup xlsx worksheet object
        :return: void. Raise IOIGeneratedError on error
        """
        # Each key is a letter', value is a list of field names. This mimics DD Wiki
        # .. (Example 'A' - see: http://ddwiki.reso.org/display/DDW/A+-+Lookup+Fields)
        top_index = self.spreadsheet_info['Lookups']
        #  Each key is a field name, value is a list of lookup values
        # .. (Example 'PropertySubType Lookups - see: http://ddwiki.reso.org/display/DDW/PropertySubType+Lookups)
        lookup_fields = {}

        header_cols = [ws.cell(row=1, column=idx).value for idx in range(1, ws.max_column+1)
                       if ws.cell(row=1, column=idx).value is not None]

        # Loop through all rows
        for row in range(2, ws.max_row + 1):
            # Each row in csv file is lookup value. Lookup field name col 'Lookup', Lookup value col 'Enumeration'
            # .. Each entry in lookup_field is a lookup field. Value is a list of lookup values
            self.fillin_lookupfield_byrow(ws, lookup_fields, header_cols, row)

        # Loop through every lookup field and create entry in top_index {}
        for fld in lookup_fields:
            if fld is not None:
                top_index.setdefault(fld[0], []).append([fld, lookup_fields[fld]])

    def fillin_lookupfield_byrow(self, ws, lookup_fields, header_cols, row):
        """ Read row from spreadsheet and translate to internal dictionary object

        :param ws: (obj) Worksheet object
        :param lookup_fields: (dict) Container for all lookup fields and values
        :param header_cols: (list) All header columns
        :param row: (int) Row being submitted
        :return: (void).  Raise IOIGeneratedError on error.
        """
        myRow = {}
        for col_num, col_val in enumerate(header_cols):
            myRow[col_val] = ws.cell(row=row, column=col_num+1).value
        try:
            lookup_fields.setdefault(myRow['LookupField'], []).append(myRow)
        except KeyError:
            raise IOIGeneratedError("[IOI-10] Cannot find field 'LookupField' in spreadsheet")
        # print("Row {}, LookupField:{}, LookupValue{}".format(row, myRow['LookupField'], myRow['LookupValue']))

    def _create_resource_dict(self, resource_name, ws):
        """ Populate xlsx resource/collection rows into internal dictionary (self.spreadsheet_info['Resources'])

        :param ws: (obj) lookup xlsx worksheet object
        :return: void. Raise IOIGeneratedError on error
        """
        self.spreadsheet_info['Resources'][resource_name] = {}

        header_cols = [ws.cell(row=1, column=idx).value for idx in range(1, ws.max_column+1)
                       if ws.cell(row=1, column=idx).value is not None]

        # Loop through all rows
        for row in range(2, ws.max_row + 1):
            myRow = {}
            if ws.cell(row=row, column=1).value is not None and len(ws.cell(row=row, column=1).value) > 0:
                for col_num, col_val in enumerate(header_cols):
                    myRow[col_val] = ws.cell(row=row, column=col_num + 1).value
                self._replace_val_in_groups(myRow)
                self.spreadsheet_info['Resources'][resource_name].setdefault(myRow["Groups"], []).append(myRow)

    def _replace_val_in_groups(self, myRow):
        """ Insert '_' char between each field in Groups column for better tree support. Assume groups separated by ','
        
        :param myRow: (obj) A row from the input .xlsx file 
        :return: Null
        """
        # Change: Added 4/17/17 to allow '_' and ',' as group separators
        if 'Groups' in myRow:  # ignore row if no 'Groups' column
            if '_' not in myRow['Groups']: # if any '_' found, assume field already separated with '_'
                myRow['Groups'] = '_' + myRow['Groups'].replace(',','_')
            else:
                if myRow['Groups'][0] != '_':
                    myRow['Groups'] = '_' + myRow['Groups']

def main(argv):
    # https://docs.python.org/3.3/library/argparse.html
    # https://docs.python.org/3/howto/argparse.html
    parser = argparse.ArgumentParser(description='RESO xlsx to xml Import')
    parser.add_argument('-f', '--working_folder', default=None,
                        help="Default folder for config, ini and error log files <current folder>")
    parser.add_argument('-x', '--xlsx_filename', default=None,
                        help="Input .xlsx file containing new fields and lookups")
    parser.add_argument('-i', '--max_id_filename', default='stat_warning_log.txt',
                        help="Input file containing DD Wiki max record and lookup ids <stat_warning_log.txt>")
    parser.add_argument('-w', '--ddwiki_export_filename', default=None,
                        help="Input Previous Confluence DD Wiki exported XML file to check for duplicate page titles")
    # type=valid_date .. validate date input with valid_date()
    parser.add_argument('-d', '--xlsx_date', type=valid_date, default=None,
                        help="Noted create date for input spreadsheet (YYYY-MM-DD)")
    parser.add_argument('-e', '--error_logging', default=20,
                        help="Error Logging Level (0-None, 10-Debug, 20-Info, 30-Warn, 40-Err, 50-Critical <20>")
    args = parser.parse_args()

    if args.working_folder is None:
        working_folder = os.getcwd()
    else:
        working_folder = args.working_folder
    input_folder = os.path.join(working_folder, "files")
    # xlsx create date (in xml root node) - Enter manually based on DD Spreadsheet
    if args.xlsx_date is None:
        xlsx_date = datetime.datetime.now()
    else:
        xlsx_date = args.xlsx_date
    logger = logging.getLogger(__project__)
    logger.setLevel(args.error_logging)
    formatter = logging.Formatter('%(asctime)s (%(name)s:%(levelname)s) %(message)s')
    stream_display = logging.StreamHandler()
    stream_display.setFormatter(formatter)
    stream_file = logging.FileHandler(os.path.join(working_folder, "logging.txt"), mode='w')
    stream_file.setFormatter(formatter)
    logger.addHandler(stream_display)
    logger.addHandler(stream_file)
    logger.info("Starting IOI xlsx-to-xml. Program verson:{0}".format(__version_date__))

    input_xlsx_filepath = os.path.join(input_folder, args.xlsx_filename)
    # max_id_filename - file created by RESOExporter. Contains max rec/lookup ids (i.e ddwiki_stat_log2017-05-12.txt)
    max_id_filepath = os.path.join(input_folder, args.max_id_filename)
    ddwiki_export_filepath = os.path.join(input_folder, args.ddwiki_export_filename)

    # Read in csv files and convert to internal python dict {}
    logger.info("Importing file:'{}' with date:{}".format(args.xlsx_filename,
                                xlsx_date.strftime('%m-%d-%Y %H:%M')))
    logger.info("Base Folder for config files is: " + working_folder)
    logger.info("Data File Folder is: " + input_folder)
    logger.info("Input RESO Export XML file:{};  Input RESO Stat/Max ID File:{}".format(args.max_id_filename,
                                                                                        args.ddwiki_export_filename))
    try:
        # Create object to convert xlsx into xml
        xlsx_to_dict = ResoXLSXtoDict(working_folder=working_folder,
                                      xlsx_filepath=input_xlsx_filepath)
    except IOIGeneratedError as e:
        logger.error("? Error initiating ResoXLSXtoDict: " + e.value)
        sys.exit(-1)
    logger.info("Resultant Output IOI XML File: " + xlsx_to_dict.xml_filename)

    try:
        # Read xlsx into internal structure xlsx_to_dict.spreadsheet_info
        xlsx_to_dict.read_xlsx_file()
    except IOIGeneratedError as e:
        logger.error("Error reading .xlsx file: " + e.value)
        sys.exit(-1)

    try:
        # Convert internal structure into IOI xml file
        xml_result = DictToXML(working_directory=working_folder,
                               result_xml_filepath=xlsx_to_dict.xml_filepath,
                               max_id_filepath=max_id_filepath,
                               ddwiki_export_filepath=ddwiki_export_filepath,
                               spreadsheet_dict=xlsx_to_dict.spreadsheet_info,
                               xlsx_date=xlsx_date,
                               program_config_data=xlsx_to_dict.config)
    except DXMLGeneratedError as e:
        logger.error("Error creating XML File: " + e.value)
        sys.exit(-1)

    logger.info('** Program Ends in Success **')

if __name__ == "__main__":
    main(sys.argv[1:])

# todo: ? Reference col in Lookups created bad link, ' Resource' not concatenated (i.e. Bot. IntTracking sb IntTracking Resource)
# ... See References column in DRAFT lookup page 'Agent (NotedBy)'. Collection had to be concatenated (https://goo.gl/0pj92s)

