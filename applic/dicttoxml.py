import datetime
import logging
import ntpath
import os
from operator import itemgetter

from lxml import etree as xml_tree
from namedentities import *
from treelib import Tree
from treelib.tree import NodeIDAbsentError, MultipleRootError

__project__ = 'IOI_Import'
__author__ = "Robert Gottesman"
__version_date__ = "04/27/2018"
__high_err_num__ = 47

""" Change log
3/30/2017 - See section: elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUP_FLDID:
3/30/2017 - The fieldname  <<LookupFieldID>> in template LookupFieldTemplate was changed to <<Lookup_FieldID>>
4/14/2017 - Fixed bug in config.ini processing such as: formatted as: [sheet tab name] = [output xml resource name]
4/17/2017 - Group column items split with ',' or '_' (previous '_' only)
5/02/2017 - Added logic for Repeating Elements / Collections (PARSE_FLD_REFERENCES)
5/09/2017 - Added logic for 'Collection' type resources. Relies on self.page_links[]
5/23/2017 - Added fnx _add_reference_sub_nodes() to handle collection field 'References' (see: PARSE_FLD_REFERENCES)
5/24/2017 - For a collection node (PARSE_FLD_COLLECTION), Node value & link attribute will be the same as in config.ini
5/24/2017 - Added ' Field' to collection reference columns for clarity 
4/24/2018 - Fixed bug in finding duplicate names within lookup values
4/25/2017 - Modified how code differentiates between Property Resource, Other Resources and Collections
"""


class DXMLGeneratedError(Exception):
    """
    Handle known problems in this module passing detail information
    """
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class DictToXML:
    XML_ROOT_TAG = 'wikiimport'
    DATETIME_FORMAT = '%m/%d/%Y %H%M'
    XLSX_DATETIME_FORMAT = "%Y%m%dT%H%M"

    # Variable in config file (DDWikiImportConfig.xml) describing parsing from xlsx to xml output file
    PARSE_SIMPLE = 0            # Convert val from xlsx for XML file
    PARSE_LABEL = 1             # Create XML Label nodes
    PARSE_PANEL = 2             # Grab info from Confluence panel (not used)
    PARSE_DATETIME = 3          # Parse Date field from xlsx to xml output file
    PARSE_LKP_PROP_REFERENCES = 4    # Parse Lookup Reference Field
    PARSE_GROUPS = 5            # Parse Resource Groups fields (with Link)
    PARSE_LOOKUP = 6            # Parse Resource Lookup field (with Link)
    PARSE_LOOKUP_STATUS = 7     # Correct bad data in xlsx - force n/a for non-lookups
    PARSE_LOOKUP_FIELD = 8      # within a Lookup Value
    PARSE_MOD_TIMESTAMP = 9     # Insert today's date (e.g Revision Date)
    PARSE_LOOKUPID = 10         # Compute LookupID
    PARSE_LOOKUP_FLDID = 11     # Compute LookupFieldID
    PARSE_RECORDID = 12         # Compute RecordID
    PARSE_FLD_REFERENCES = 13   # Parse Reference columns in Resource (collections)
    PARSE_FLD_COLLECTION = 14   # Parse a 'Collection' column to point to collection resource

    IGNORE_FIELDS = ['OriginalEntryTimestamp']
    INTERNAL_OUTPUT_DATE_FORMAT = '%Y%m%dT%H%M'
    DEFAULT_DATE_FORMAT = '%Y%m%d'
    PROP_NOLOOKUP_TEMPLATE = 'PropNoLookupResourceTemplate'
    OTHER_NOLOOKUP_TEMPLATE = 'OtherNoLookupResourceTemplate'
    REFERENCE_NOLOOKUP_TEMPLATE = 'ReferenceNoLookupResourceTemplate'
    STANDARD_NAME_COLUMN = 'StandardName'
    SPECIAL_PAGE_SUFFIX = ['Resource', 'Group', 'Collection', 'Fields', 'Values', 'Lookups']

    def __init__(self, files_and_folders, max_id_filepath, ddwiki_exported_filepath,
                 result_xml_filepath,
                 spreadsheet_dict,
                 xlsx_date,
                 program_config_data=None):
        """ Convert internal .xlsx dict to specially formatted XML file to be used for importing into Confluence DD Wiki

        :param files_and_folders: (obj) object containing file locations
        :param max_id_filepath: (str) File name/path for file containing DD Wiki max lookupids (stat_warning_log.txt)
        :param ddwiki_exported_filepath: (str) File name/path for latest dd wiki xml exported file
        :param result_xml_filepath: (str) File name/path for resultant IOI import xml file
        :param spreadsheet_dict: (dict) xlsx file data converted into internal dict format
        :param xlsx_date: (datetime) Timestamp for result_xml_filepath
        :param program_config_data: (dict) config.ini file read into dictionary
        :return: None. Raise DXMLGeneratedError on error.
        """
        self.logger = logging.getLogger(__project__ + '.' + self.__class__.__name__)
        self.report_warning = True  # Report certain warning messages only once
        self.start_datetime = datetime.datetime.today()
        self.date_format_notime = '%b %d %Y'
        self.date_format_withtime = '%b %d %Y %I:%M %p'  # Uses AM/PM format
        self.start_datetime_str = self.start_datetime.strftime(self.date_format_notime)
        self.spreadsheet_data = spreadsheet_dict    # xlsx converted into a dictionary
        self.field_and_lookup_names = []   # Used to ensure unique page titles
        self.program_config_data = program_config_data  # setup info from config.ini
        self.resource_descriptions = {}  # retrieved from config.ini
        self.page_links = {}  # Translate xlsx columns into appropriate text for Lookup page links (config.ini)
        self.resource_tree = None  # Create internal tree for Wiki output structure (xml output file)
        self._read_ini_config_data()  # Convert config.ini info into dict {}
        self.xml_config_data = self._read_xml_config_file(files_and_folders)  # Read config. defines xlsx->xml rules
        self._read_max_ids(max_id_filepath)
        self._load_page_titles_from_ddwiki_export(ddwiki_exported_filepath)  # check for dup confluence page titles
        self.xml_root = xml_tree.Element(self.XML_ROOT_TAG)  # Setup root output XML node
        self.xml_root.set('XMLCreateDate', self.start_datetime.strftime(self.INTERNAL_OUTPUT_DATE_FORMAT))
        self.xml_root.set('XlsxDate', xlsx_date.strftime(self.INTERNAL_OUTPUT_DATE_FORMAT))
        # Populate output xml structure .. the write file out
        self._create_resources()  # Create resource and collection nodes in IOI import xml
        self._create_lookups()  # Create lookup fields/value nodes in IOI import xml
        self.write_xml_file(result_xml_filepath)

    def _load_page_titles_from_ddwiki_export(self, ddwiki_exported_filepath):
        """ Load Page Titles from exported xml file. Needed to check for duplicate Confluence page titles.
        .. store into variable: (dict) field_and_lookup_names

        :param ddwiki_exported_filepath: (str) File name/path for latest dd wiki xml exported file
        :return: None. Raise DXMLGeneratedError on error.
        """
        if not os.path.exists(ddwiki_exported_filepath):
            raise DXMLGeneratedError('[DXM-30] Cannot find pre DD Wiki exported xml file ' + ddwiki_exported_filepath)

        # http://lxml.de/api/lxml.etree.XMLParser-class.html
        xml_tree.XMLParser(remove_blank_text=True, resolve_entities=False)
        try:
            xml_root = xml_tree.parse(ddwiki_exported_filepath)
        except OSError:
            raise DXMLGeneratedError("[DXM-31] Cannot Open DD Wiki Input XML file: " + ddwiki_exported_filepath)
        root = xml_root.getroot()
        resource_nodes = root.findall(".//StandardName")
        for name_node in resource_nodes:
            if name_node.text not in self.field_and_lookup_names:
                self.field_and_lookup_names.append(name_node.text)
        lookupval_nodes = root.findall(".//LookupValue")
        for name_node in lookupval_nodes:
            if name_node.text not in self.field_and_lookup_names:
                self.field_and_lookup_names.append(name_node.text)
        self.logger.info("[DXM-33] Note on Existing DD Wiki: {} fields found, {} lookup values found".
                         format(len(resource_nodes), len(lookupval_nodes)))

    def _read_ini_config_data(self):
        """ Read [Resource-Descriptions] and [PageLinks] sections from config.ini to internal dicts {}

        :return: None
        """
        # low: Add try/except while reading config.ini
        for desc in self.program_config_data['Resource-Descriptions']:
            # Add period to end of description so it looks like a sentence
            self.resource_descriptions[desc] = self.program_config_data['Resource-Descriptions'][desc] + '. '
        for desc in self.program_config_data['PageLinks']:
            self.page_links[desc] = self.program_config_data['PageLinks'][desc]

    def _read_xml_config_file(self, files_and_folders):
        """ Read config file DDWikiImportConfig.xml which describes xlsx input format and xml output format
        .. Each xml 'Form' node has children describing the fields it can expect based on Confluence page

        :param files_and_folders: (obj) object containing file locations
        :return: {dict} representation of config file. Raise DXMLGeneratedError on error.
        """
        config_filename = os.path.join(files_and_folders.config_folder, "DDWikiImportConfig.xml")
        try:
            config_tree = xml_tree.parse(config_filename)
        except (IOError, xml_tree.XMLSyntaxError):
                raise DXMLGeneratedError('[DXM-10] Cannot Open/Process Config File ' + config_filename)

        config = {}  # Translate config xml file into internal dictionary
        config_root = config_tree.getroot()
        for ele in config_root:
            try:
                # The children nodes of Parent node 'Form' defines fields for a specific Confluence page type
                # .. The node.text is the dict key which is the field header text that appears in Confluence page.
                if ele.tag == "Form":
                    config[ele.get("Name")] = {}  # Attribute 'Name' holds 'Form:' value from Confluence page
                    config[ele.get("Name")]['Attributes'] = {'Sequence': 0}
                    for key, value in ele.items():
                        config[ele.get("Name")]['Attributes'][key] = value
                    for field in ele:
                        # Create a dictionary for each output XML field
                        # .. Sequence (XML output sequence-not used), XMLName (Output xml node tag)
                        # .. Value (value of node .. may be used in place of xlsx value
                        # .. ParsingCode (how to parse from xlsx) - 'PARSE_' options above
                        # .. ChildTagName (child name of field node if it is to have children nodes
                        # .. AutoCompute (value for XML is derived from computation and NOT the xlsx
                        config[ele.get("Name")][field.get("XMLName")] = {"Sequence": int(field.get("Sequence")),
                                                                         "Value": field.text,
                                                                         "ParsingCode":
                                                                             int(field.get("ParsingCode", int(self.PARSE_SIMPLE))),
                                                                         "ChildTagName": field.get("ChildTagName"),
                                                                         "AutoCompute": field.get("AutoCompute"),
                                                                         "CollectionTemplate": field.get(
                                                                             "CollectionTemplate"),
                                                                         "DefaultValue": field.get("DefaultValue")}
            except KeyError:
                raise DXMLGeneratedError("[DXM-01] Ill formed XML in config file: " +
                                         str(config_filename).split('\\')[-1:][0])

        return config

    def _add_date_node(self, parent_node, nodes_from_config, config_node_text, page_title, xlsx_values):
        """ Convert xlsx date into XML date format

        :param parent_node: xml node # XML node to add this date element
        :param nodes_from_config: dict # config dictionary which describes how to handle all fields
        :param config_node_text: str # specific entry (key) in nodes_from_config. This will be XML tag
        :param page_title: str # The page title containing this date field
        :param xlsx_values: dict # Values from xlsx row to be inserted in XML file
        :return: (str) Date Value added to column - or None if error. Raise DXMLGeneratedError on error
        """
        if nodes_from_config[config_node_text]['AutoCompute'] == 'Y':
            # AutoCompute in a date field only works for ModificationTimestamp
            if config_node_text == 'ModificationTimestamp':
                val = self.start_datetime_str
            else:
                raise DXMLGeneratedError("[DXM-02] Cannot resolve AutoCompute Date '{0}' for page {1}".
                                         format(config_node_text, page_title))
        else:
            # Get date from xlsx and format it for XML
            if nodes_from_config[config_node_text]['DefaultValue'] is not None:
                # string format of "YYYYMMDDTHHMM"
                val = nodes_from_config[config_node_text]['DefaultValue']
                if val == "*":  # Use today's date if entry is blank
                    default_date_str = datetime.datetime.now().strftime(self.DEFAULT_DATE_FORMAT)
                else:
                    default_date_str = val
            else:
                default_date_str = None
            dte_err = "[DXM-03] Cannot find column/xlsx_values for date field '{0}' in row/page '{1}'". \
                format(nodes_from_config[config_node_text]['Value'], page_title)
            # Check if column exists in .xlsx row
            if nodes_from_config[config_node_text]['Value'] in xlsx_values:
                dte_val = xlsx_values[nodes_from_config[config_node_text]['Value']]
                # Assign default date if xlsx_values in .xlsx is empty
                if dte_val is None or (isinstance(dte_val, str) and len(dte_val) == 0):
                    if default_date_str is not None:
                        dte_val = default_date_str
                    else:
                        # if cell empty and no default xlsx_values, error
                        raise DXMLGeneratedError(dte_err)
            else:
                if default_date_str is None:
                    raise DXMLGeneratedError(dte_err)  # Cell is missing in xlsx and no default value
                else:
                    dte_val = default_date_str  # Cell is missing in xlsx, but default value stated

            if not isinstance(dte_val, datetime.date):
                if not isinstance(dte_val, str):
                    dte_val = str(dte_val)
                try:
                    dte_val = dte_val.strip()
                    if len(dte_val) == 8:
                        dte_val += "T0000"
                    # Accepting xlsx string format of "YYYYMMDDTHHMM"
                    dte_val = datetime.datetime.strptime(dte_val, self.XLSX_DATETIME_FORMAT)
                except (ValueError, TypeError):
                    raise DXMLGeneratedError("[DXM-12] xlsx cell not in Date format. Column {0} in page {1}".
                                             format(config_node_text, page_title))
            if dte_val.hour == 0 and dte_val.minute == 0:
                val = dte_val.strftime(self.date_format_notime)
            else:
                val = dte_val.strftime(self.date_format_withtime)
        dte_node = xml_tree.SubElement(parent_node, config_node_text)
        dte_node.text = val
        return val

    def _add_label_sub_nodes(self, node_tag, sub_node_tag, sub_node_value_str):
        """ Create xml parent/child tags for Labels

        :param node_tag (str):  parent 'Label' Node tag (s.b. 'Labels')
        :param sub_node_tag (str): Child node tag names (s.b. 'Labels')
        :param sub_node_value_str (str): Child node values separated by commas
        :return parent node:
        """
        parent_node = xml_tree.Element(node_tag)
        for sub_value_str in sub_node_value_str.split(','):
            sub_node = xml_tree.SubElement(parent_node, sub_node_tag)
            sub_node.text = sub_value_str
        return parent_node

    def _add_group_sub_nodes(self, node_tag, sub_node_tag, sub_node_value_str):
        """ Create xml parent/child tags for Resource node Groups defining how pages will appear in confluence nav panel

        :param node_tag (str):  parent of Resource 'Groups' Node tag (s.b. Groupings)
        :param sub_node_tag (str): child of node_tag (s.b. 'Group')
        :param sub_node_value_str (str): Child node values separated by commas
        :return node_tag xml:
        """
        parent_node = xml_tree.Element(node_tag)
        for sub_value_str in sub_node_value_str:
            if sub_value_str in self.page_links:
                page_link = self.page_links[sub_value_str]
            else:
                page_link = sub_value_str
            sub_node = xml_tree.SubElement(parent_node, sub_node_tag, {'Link': page_link})
            sub_node.text = page_link
        return parent_node

    def _add_linked_sub_nodes(self, node_tag, sub_node_tag, tag_value, page_title):
        """ Translate xlsx columns into appropriate text for xml Lookup page links

        :param node_tag (str): parent of Resource Property Types Node tag (s.b. Groupings)
        :param sub_node_tag (str): child tag (s.b. 'Class)
        :param sub_node_value_str (str): xlsx column names separated by commas
        :param xlsx_values (str): Child node values separated by comma
        :param page_title (str): Current ddwiki page being processed (needed for error reporting)
        :return: None. Raise DXMLGeneratedError on error
        """
        parent_node = xml_tree.Element(node_tag)
        if tag_value is None:
            raise DXMLGeneratedError("[DXM-45] Found Null/Empty Value for Reference within column '{}' on page '{}'".
                                     format(parent_node.tag, page_title))
        for ref_text in tag_value.replace(' ', '').split(','):
            try:
                new_node = xml_tree.SubElement(parent_node, sub_node_tag, {'Link': self.page_links[ref_text]})
            except KeyError:
                err_msg = "[DXM-14] Cannot create link for ref '{}' within column '{}' on page '{}'. " \
                          "Check section PageLinks in config.ini"
                raise DXMLGeneratedError(err_msg.format(ref_text, parent_node.tag, page_title))
            new_node.text = ref_text
        return parent_node

    def _add_reference_sub_nodes(self, node_tag, sub_node_tag, tag_value):
        """ Translate xlsx columns into appropriate text for xml reference tag page links

        :param node_tag (str): parent of Resource Property Types Node tag (s.b. Groupings)
        :param sub_node_tag (str): child tag (s.b. 'Class)
        :param sub_node_value_str (str): xlsx column names separated by commas
        :param xlsx_values (str): Child node values separated by comma
        :return: None. Raise DXMLGeneratedError on error
        """
        parent_node = xml_tree.Element(node_tag)
        if tag_value is None:
            raise DXMLGeneratedError("[DXM-37] Found Null/Empty Value for Reference within {}".format(parent_node.tag))
        for ref_text in tag_value.replace(' ', '').split(','):
            # low: Reference column values that are not unique page names will not work.
            new_node = xml_tree.SubElement(parent_node, sub_node_tag, {'Link': ref_text + ' Field'})
            # Adding ' Field' to visual field for clarity
            new_node.text = ref_text + ' Field'
        return parent_node

    def _sort_nodes(self, config_element):
        """ Sort so metadata elements (within a page) appear always in same order

        :param config_element: The form name (format) that dictate the fields being sorted
        :return: list sorted list. Each element in is a 2 entry list [[0]=sort#; [1]=fieldname]
        """
        sorted_list = [[config_element[fieldname]["Sequence"], fieldname] for fieldname in config_element]
        sorted_list.sort()
        return sorted_list

    def _make_page_title(self, full_page_title, dup_qualifier='', page_template=None):
        """ Check proposed page title for duplicate .. if so append resource or lookup field as addl qualifier

        :param full_page_title (str): Proposed full page title name
        :param dup_qualifier (str): Resource or Field Lookup name to add to title
        :return proposed page title (str):
        """
        # Field names have ' Field' concatenated to end. Remove text

        if len(full_page_title.rsplit(' ', 1)) == 1:
            suffix = full_page_title
        else:
            try:
                suffix = full_page_title.rsplit(' ', 1)[1]  # Get last word in title
            except IndexError:
                raise DXMLGeneratedError("[DXM-44] Index Error creating page title with '{}'".format(full_page_title))
        if suffix in self.SPECIAL_PAGE_SUFFIX:  # Don't check special page names
            return full_page_title
        if page_template == 'LookupValueTemplate':
            item_name = full_page_title
            suffix = ''
        else:
            item_name = full_page_title.split(' ')[0]  # Get 1st word
        # See if the Name exists
        if item_name in self.field_and_lookup_names:
            page_title = item_name + ' (' + dup_qualifier + ') ' + suffix
        else:
            self.field_and_lookup_names.append(item_name)
            page_title = full_page_title
        return page_title

    def _adjust_resource_page_template(self, this_node):
        """ Non lookups fields use page template PropNoLookupResourceTemplate or
            OtherNoLookupResourceTemplate

        :param this_node (xml node): Node that will adjust attribute
        :return: None
        """
        # A collection cannot have a lookup
        if 'Collection' in this_node.attrib['Page_Template']:  # Collection pages have same format
            return
        if 'Prop' == this_node.attrib['Page_Template'][0:4]:
            this_node.attrib['Page_Template'] = self.PROP_NOLOOKUP_TEMPLATE
        elif 'Reference' in this_node.attrib['Page_Template']:
            this_node.attrib['Page_Template'] = self.REFERENCE_NOLOOKUP_TEMPLATE
        else:
            this_node.attrib['Page_Template'] = self.OTHER_NOLOOKUP_TEMPLATE

    def _add_xml_nodes(self, parent_node, nodes_from_config, value='', other_page_title=None,
                       replace_labels=None, resource_name=''):
        """ Add nodes to xml output based on xml config file (DDWikiImportConfig.xml)

        :param parent_node (xml node): XML node to add this date element
        :param nodes_from_config (dict): config dictionary which describes how to handle all fields
        :param value (str or dict): Value (autocompute - str) or from xlsx (dict)
        :param other_page_title (str): Preferred Page Title
        :param replace_labels (str): optional labels separated by commas
        :param resource_name: optional String used to make page title unique
        :return (xml node): Node added to XML structure and children. Raise DXMLGeneratedError on error
        """
        if other_page_title is None:
            page_title = nodes_from_config['Attributes']['Page_Title'].strip()
        else:
            page_title = other_page_title.strip()
        # todo: Possibly create attribute after loop. Page_Template might change within loop (need sample case for this)
        attrs = {'Page_Template': nodes_from_config['Attributes']['Page_Template'],
                 'Page_Title': self._make_page_title(page_title, resource_name,
                                                     nodes_from_config['Attributes']['Page_Template'])}
        prime_node = xml_tree.SubElement(parent_node, nodes_from_config['Attributes']['Node_Type'], attrib=attrs)
        sorted_nodes = self._sort_nodes(nodes_from_config)  # Sort by nodes_from_config Sequence attribute
        # Loop through xml nodes in config file DDWikiImportConfig.xml to create final IOI xml nodes for output
        for config_node_list in sorted_nodes:
            config_node_text = config_node_list[1]  # value of XMLName attribute from DDWikiImportConfig.xml
            # Nodes attributes from XML are placed into the dict{} (attributes already added above)
            if config_node_text != 'Attributes' and config_node_text not in self.IGNORE_FIELDS:
                if nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LABEL:  # (1) Parse Child Nodes
                    # labels in a str separated by commas
                    if replace_labels is None:
                        lbls = nodes_from_config[config_node_text]['Value']
                    else:
                        lbls = replace_labels
                    prime_node.append(self._add_label_sub_nodes(node_tag=config_node_text,
                                                                sub_node_tag=nodes_from_config[config_node_text][
                                                                    'ChildTagName'],
                                                                sub_node_value_str=lbls))
                # (0) Grab value from xlsx dict
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_SIMPLE:
                    # val = page_title.replace(' ', '_').lower()
                    if nodes_from_config[config_node_text]['AutoCompute'] == 'Y':
                        if config_node_text == 'lookupfield_ref':
                            val = page_title.replace(' ', '_').lower()
                        elif config_node_text == 'Resource_Description':
                            try:
                                # Add ' Collection to end for Collections only. Otherwise considered resource
                                if ' Collection' in other_page_title:
                                    val = self.resource_descriptions[other_page_title]
                                else:
                                    val = self.resource_descriptions[other_page_title.split(' ')[0]]
                            except KeyError:
                                raise DXMLGeneratedError("[DXM-28] Cannot find resource description for '{}'".
                                                         format(resource_name))
                        else:
                            raise DXMLGeneratedError("[DXM-04] Cannot resolve AutoCompute for col '{0}' in page '{1}".
                                                     format(config_node_text, page_title))
                    else:
                        # Get value from xlsx (as translated into dict)
                        if type(value) == dict:
                            try:  # Get value from xlsx
                                val = value[nodes_from_config[config_node_text]['Value']]
                                # If no entry for this col, then use DefaultValue if one is entered
                                if val is None and nodes_from_config[config_node_text]['DefaultValue'] is not None:
                                    val = nodes_from_config[config_node_text]['DefaultValue']
                            except KeyError:
                                if nodes_from_config[config_node_text]['DefaultValue'] is not None:
                                    # Column does not exist in .xlsx, use DefaultValue
                                    val = nodes_from_config[config_node_text]['DefaultValue']
                                else:
                                    raise DXMLGeneratedError("[DXM-05] Missing column '{0}' in row/page {1}".
                                                             format(config_node_text, page_title))
                        else:
                            val = value  # Value as passed as parameter
                    new_node = xml_tree.SubElement(prime_node, config_node_text)
                    if val is None:
                        new_node.text = ''
                    elif isinstance(val, str):
                        new_node.text = val.replace('&#13;', ' ')
                    else:
                        new_node.text = str(val)

                    # new_node.text = entities(val, 'hex') # Convert special chars to XML hex node
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_DATETIME:   # (3) Date Fields
                    # Add date field and store in prime_node
                    self._add_date_node(prime_node, nodes_from_config, config_node_text, page_title, value)
                # (4) Lookup 'References' Field with Links
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LKP_PROP_REFERENCES:
                    if nodes_from_config[config_node_text]['Value'] not in value:
                        raise DXMLGeneratedError(
                            "[DXM-40] Unable to Find Column '{}' in spreadsheet for page:{}".
                                format(nodes_from_config[config_node_text]['Value'], page_title))
                    prime_node.append(self._add_linked_sub_nodes(node_tag=config_node_text,
                                                    sub_node_tag=nodes_from_config[config_node_text]['ChildTagName'],
                                                    tag_value=value[nodes_from_config[config_node_text]['Value']],
                                                                 page_title=page_title))
                # (5) Resource field 'Group' with Links
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_GROUPS:
                    prime_node.append(self._add_group_sub_nodes(config_node_text,
                                                                nodes_from_config[config_node_text]['ChildTagName'],
                                                                value[nodes_from_config[config_node_text]['Value']]))
                # (6) Resource 'Lookup' Field
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUP:
                    # xlsx correction - Force 'n/a' for non lookup fields
                    attrib = None
                    try:
                        val = value[nodes_from_config[config_node_text]['Value']]
                    except KeyError:
                        raise DXMLGeneratedError("[DXM-38] Unable to Find Column '{}' in spreadsheet for page:{}".
                                                 format(config_node_text, page_title))

                    # The word 'List' in Simple Data Type means we have a Lookup field (i.e. String List, Single)
                    if 'List' in value[nodes_from_config['Simple_Data_Type']['Value']]:
                        # Possible comment in Lookup field signaled by '<'
                        if val is None or len(val) == 0:
                            val = '<Not Defined>'
                        if len(val) > 0 and val[0] != '<':
                            if val[-8:] != ' Lookups':
                                val += ' Lookups'       # Lookup field is title + ' Lookups'
                            attrib = {'Link': val}
                        if val[0] == '<':           # Do NOT use lookup template when comment present
                            atr = prime_node.attrib['Page_Template']
                            # Not sure if this logic is used anymore
                            atr_idx = atr.find('Resource')
                            if atr_idx < 0:
                                raise DXMLGeneratedError("[DXM-32] Expecting text 'Resource' in col '{}' field '{}' ".
                                                         format(config_node_text, page_title))
                            else:
                                prime_node.attrib['Page_Template'] = atr[0:atr_idx] + 'NoLookup' + atr[atr_idx:]
                    else:
                        # Non lookups fields use page template PropNoLookupResourceTemplate or
                        # ... OtherNoLookupResourceTemplate
                        self._adjust_resource_page_template(this_node=prime_node)
                        if val != '<n/a>' and val is not None:
                            self.logger.warning("[DXM-41] Lookup Value should be n/a for col '{}' on page '{}'" 
                                                "in resource {} due to SimpleDataType".
                                                format(config_node_text, page_title, resource_name))
                        val = '<n/a>'  # Force n/a for non lookups w/no comments
                    new_node = xml_tree.SubElement(prime_node, config_node_text, attrib)
                    new_node.text = entities(val, 'hex')
                # (7) Resource 'Lookup_Status' Field
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUP_STATUS:
                    try:
                        val = value[nodes_from_config[config_node_text]['Value']]
                    except KeyError:
                        raise DXMLGeneratedError("[DXM-17] Cannot find column '{}' in field '{}'".
                                                 format(nodes_from_config[config_node_text]['Value'], page_title))
                    # xlsx correction - Force 'n/a' for non lookup fields
                    if 'List' not in value[nodes_from_config['Simple_Data_Type']['Value']]:
                        val = '<n/a>'
                    elif val is None or len(val) == 0:
                        val = '<Not Defined>'
                    new_node = xml_tree.SubElement(prime_node, config_node_text)
                    new_node.text = entities(val, 'hex')
                # (8) lookup field within a lookup value page
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUP_FIELD:
                    try:
                        val = value[nodes_from_config[config_node_text]['Value']]
                    except:
                        raise DXMLGeneratedError(
                            "[DXM-35] Cannot find value for '{}' column. Looking at node {} in page {} ".
                                format(nodes_from_config[config_node_text]['Value'], config_node_text, page_title))
                    attrib = None
                    if val is not None and len(val) > 0 and val[0] != '<':
                        # If last 8 chars has lookups .. then no need to add
                        if val[-8:] != ' Lookups':
                            val += ' Lookups'
                        attrib = {'Link': val}
                    new_node = xml_tree.SubElement(prime_node, config_node_text, attrib)
                    new_node.text = val
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUPID:
                    lookup_field_name = value[nodes_from_config['Lookup_Field']['Value']]
                    # If LookupID is not in xlsx then it will be computed
                    try:
                        lookupid_value = value[nodes_from_config[config_node_text]['Value']]
                    except KeyError:
                        lookupid_value = None
                    if lookupid_value is None or len(lookupid_value) == 0:
                        lookupid_value = self._compute_lookupid(lookup_field_name=lookup_field_name)
                        if lookupid_value < 0:
                            raise DXMLGeneratedError("[DXM-15] Program can only accomodate max 999 lookup values {}:{}".
                                        format(lookup_field_name, value[nodes_from_config['Lookup_Value']['Value']]))
                    new_node = xml_tree.SubElement(prime_node, config_node_text)
                    new_node.text = str(lookupid_value)
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_LOOKUP_FLDID:
                    # (3/30/2017) This can be reached from a lookup field and a lookup value. column names differ
                    # (3/30/2017) lookup_fieldid_value MUST be computed and not taken from spreadsheet
                    lookup_fieldid_value = None
                    if attrs['Page_Template'] == 'LookupFieldTemplate':
                        lookup_field_name = attrs['Page_Title'].replace(' Lookups', '')
                    else:
                        lookup_field_name = value[nodes_from_config['Lookup_Field']['Value']]
                        # lookup_fieldid_value = value[nodes_from_config[config_node_text]['Value']]
                    if lookup_fieldid_value is None or len(lookup_fieldid_value) == 0:
                        lookup_fieldid_value = self._compute_lookup_fieldid(lookup_field_name=lookup_field_name)
                    new_node = xml_tree.SubElement(prime_node, config_node_text)
                    new_node.text = str(lookup_fieldid_value)
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_RECORDID:
                    # If field_name is not in xlsx then it will be computed
                    try:
                        recordid_value = value[nodes_from_config[config_node_text]['Value']]
                    except KeyError:
                        recordid_value = None
                    if recordid_value is None or len(recordid_value) == 0:
                        recordid_value = self._compute_recordid(resource_name=resource_name)
                        if recordid_value < 0:
                            raise DXMLGeneratedError("[DXM-18] Program has max 999 recordid values for resource {}".
                                                     format(resource_name))
                    new_node = xml_tree.SubElement(prime_node, config_node_text)
                    new_node.text = str(recordid_value)
                # (13) Parse Reference columns in Resource (collections)
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_FLD_REFERENCES:
                    prime_node.append(self._add_reference_sub_nodes(node_tag=config_node_text,
                                                    sub_node_tag=nodes_from_config[config_node_text]['ChildTagName'],
                                                    tag_value=value[nodes_from_config[config_node_text]['Value']]))
                    # xlsx correction - Force 'n/a' for non lookup fields
                elif nodes_from_config[config_node_text]['ParsingCode'] == self.PARSE_FLD_COLLECTION:
                    try:
                        val = value[nodes_from_config[config_node_text]['Value']]
                    except KeyError:
                        raise DXMLGeneratedError(
                            "[DXM-46] Cannot find 'Collection' column for page '{}'".format(page_title))
                    # Remove word collection if present in xlsx so we match w/key in config.ini
                    if val is not None:
                        val = val.replace(' Collection', '')
                        if val not in self.page_links:
                            raise DXMLGeneratedError(
                                "[DXM-36] Cannot create link for collection column value '{}' within col {}. "
                                "Check section PageLinks in config.ini".
                                format(val, parent_node.tag))
                        attrib = {'Link': self.page_links[val]}
                        # Replace prime_node['Page_Template'] with correct Collection named template
                        try:
                            prime_node.attrib['Page_Template'] = \
                                nodes_from_config[config_node_text]['CollectionTemplate']
                        except KeyError:
                            raise DXMLGeneratedError("[DXM-39] No 'CollectionTemplate' attr in DDWikiImportConfig {}-{}"
                                                     .format(val, parent_node.tag))
                        new_node = xml_tree.SubElement(prime_node, config_node_text, attrib)
                        # The node value and Link attribute are the same (5/24/2017)
                        new_node.text = self.page_links[val]
                else:
                    # self.logger.warning("[E100] No logic for parse code {0} on page '{1} ".
                    #            format(nodes_from_config[config_node_text]['ParsingCode'], page_title))
                    if self.report_warning:
                        self.logger.warning("[DXM-06] No Program Code for {0} in page {1}".
                                        format(config_node_text, page_title))
                        self.report_warning = False
        return prime_node

    def _compute_lookupid(self, lookup_field_name):
        """ Compute max lookupid for the lookup value.  Each lookupfield id is incremented by 1000.
        .. Hence (in this program) the max # of lookup values in a lookup field is 1000

        :param lookup_field_name: (str) Lookup Field name to determine max lookup id for that field
        :return: (int) Highest lookupid for tht field.
        """
        # If the lookup field doesn't already exist create new lookup field id#
        self._add_lookup_fieldid(lookup_field_name)

        if (self.max_lookupids[lookup_field_name] + 1) % 1000 > 998:
            return -1  # Program can only accomodate max 999 lookup values
        self.max_lookupids[lookup_field_name] += 1
        return self.max_lookupids[lookup_field_name]

    def _compute_lookup_fieldid(self, lookup_field_name):
        """ Compute the lookup field id for given lookup field

        :param lookup_field_name:
        :return: (int) LookupFieldID. Raise DXMLGeneratedError on error
        """
        # If the lookup field doesn't already exist create new lookup field id#
        self._add_lookup_fieldid(lookup_field_name)
        try:
            return self.max_lookupids[lookup_field_name] - (self.max_lookupids[lookup_field_name] % 1000)
        except TypeError:
            print("? Issue with " + lookup_field_name)
            raise DXMLGeneratedError("[DXM-25] Computing max field lookupid issue witn {}.".format(lookup_field_name))

    def _compute_recordid(self, resource_name):
        """ Compute max recordid for the lookup value.  Each lookupfield id is incremented by 1000.
        .. Hence (in this program) the max # of lookup values in a lookup field is 1000

        :param resource_name: (str) Field name to determine max lookup id for that field
        :return: (int) Highest lookupid for tht field. Raise DXMLGeneratedError on error
        """
        # If the lookup field doesn't already exist create new lookup field id#
        self._add_recordid(resource_name)

        if (self.max_recordids[resource_name] + 1) % 1000 > 999:
            raise DXMLGeneratedError("[DXM-29] This program only allows a max of 999 fields in resouce {}.".
                                     format(resource_name))

        self.max_recordids[resource_name] += 1
        return self.max_recordids[resource_name]

    def _add_lookup_fieldid(self, lookup_field_name):
        """ If the lookup field doesn't already exist create new lookup field id#

        :param lookup_field_name: (str) lookup field name
        :return: (int) LookupFieldID
        """
        if lookup_field_name not in self.max_lookupids:
            # If the lookup field doesn't already exist create new lookup field id#
            self.max_id = self.max_id - (self.max_id % 1000) + 1000
            self.max_lookupids[lookup_field_name] = self.max_id
            self.logger.info(
                "[DXM-43] Creating Base Max Lookup Field ID for '{}' with value {}".
                    format(lookup_field_name, self.max_id))

    def _add_recordid(self, resource_name):
        """ If the resource name doesn't already exist create new field id#

        :param resource_name: (str) lookup field name
        :return: (int) RecordID
        """
        if resource_name not in self.max_recordids:
            # If the lookup field doesn't already exist create new lookup field id#
            self.max_id = self.max_id - (self.max_id % 1000) + 1000
            self.max_recordids[resource_name] = self.max_id
            self.logger.info(
                "[DXM-42] Creating Base Max RecordID for resource '{}' with value {}".
                    format(resource_name, self.max_id))

    def _create_resource_nodes(self, parent_xml_node, sheet_tab_name, resource_name,
                               group_name, level_key, config_form_name):
        """ Build XML Resource nodes. Tree structure of nodes are laid out in self.resource_tree

        :param parent_xml_node: Parent node for new node created
        :param sheet_tab_name: tab name in xlsx that represents resource
        :param resource_name (str): Name of top resource as it appears in output xml
        :param group_name (str): Name of group node (top resource node has tag name of 'Group')
        :param level_key (str): Where node belongs from xlsx 'Group' column. (Levels separated by '_'?)
        :param config_form_name (str): Name attribute value in DDWikiImportConfig.xml which defines fields that apply
        :return: None. Raise DXMLGeneratedError on error
        """
        # Get info on Node. This is a treelib variable
        this_node = self.resource_tree.get_node(level_key)
        # Top level is Resource lower levels are Group
        if self.resource_tree.depth(this_node) == 0:
            node_type = 'Resource'
            if group_name in self.page_links:
                page_title = self.page_links[group_name]
            else:
                page_title = \
                    self.xml_config_data["Resource"]['Attributes']['Page_Title'].replace('[[Name]]', group_name)
        else:
            node_type = 'Group'
            page_title = \
                self.xml_config_data["Group"]['Attributes']['Page_Title'].replace('[[Name]]', group_name)
        # Add Group or Resource Node (Items are underneath)
        this_level_xml_node = self._add_xml_nodes(parent_node=parent_xml_node,
                                                  nodes_from_config=self.xml_config_data[node_type],
                                                  other_page_title=page_title,
                                                  resource_name=resource_name)
        # Add item nodes underneath Group or Resource node as added above
        if level_key in self.spreadsheet_data['Resources'][sheet_tab_name]:
            # self.spreadsheet_data['Resources'][resource_name][level_key] returns a list [] of dict items
            # itemgetter() := Return a callable object that fetches item from its operand using  __getitem__()
            sorted_item_nodes = sorted(self.spreadsheet_data['Resources'][sheet_tab_name][level_key],
                             key=itemgetter(self.STANDARD_NAME_COLUMN))
            for item_node in sorted_item_nodes:
                item_name = item_node[self.STANDARD_NAME_COLUMN]
                page_title = \
                    self.xml_config_data[config_form_name]['Attributes']['Page_Title'].replace('[[Name]]', item_name)
                new_node = self._add_xml_nodes(this_level_xml_node,
                                               nodes_from_config=self.xml_config_data[config_form_name],
                                               other_page_title=page_title,
                                               value=item_node,
                                               resource_name=resource_name)
                # Add additional labels
                labels_node = new_node.find('Labels')
                extra_labels = []
                # Need to create a label for a Multi/Single select (lookup) field
                lkup_node = new_node.find('Lookup')
                if lkup_node is not None and len(lkup_node.text) > 0 and 'Link' in lkup_node.attrib:
                    extra_labels.append(lkup_node.attrib['Link'].replace(' ', '_').lower())
                prop_node = new_node.find('Property_Types')
                # Create labels for each property class applied to this field
                if prop_node is not None:
                    for cls in prop_node.findall('Class'):
                        extra_labels.append('prop_' + cls.text)
                for lbl in extra_labels:
                    label_node = xml_tree.SubElement(labels_node, 'Label')
                    label_node.text = lbl.lower()
        # Add Children Items and Groups (recursion)
        try:
            child_nodes = self.resource_tree.children(level_key)
        except NodeIDAbsentError:
            raise DXMLGeneratedError("[DXM-07] Unable to find value '{0}' in xlsx 'Group' column for resource '{1}'".
                                     format(level_key, resource_name))

        for child_node in child_nodes:
            self._create_resource_nodes(parent_xml_node=this_level_xml_node,
                                        sheet_tab_name=sheet_tab_name,
                                        resource_name=resource_name,
                                        group_name=child_node.tag,
                                        level_key=level_key + ',' + child_node.tag,
                                        config_form_name= config_form_name)

    def _build_resource_tree(self, sheet_tab_name):
        """ Create a tree structure (using treelib) for resource to mimic final output DD Wiki xml structure

        :param sheet_tab_name (str):
        :return: None. Raise DXMLGeneratedError on error
        """
        self.resource_tree = Tree()  # Tree() from treelib
        try:
            resource_levels = sorted(self.spreadsheet_data['Resources'][sheet_tab_name].keys())
        except (TypeError, KeyError):
            raise DXMLGeneratedError("[DXM-27] Spreadsheet tab '{}' has blank lines or is unstructured".
                                     format(sheet_tab_name))

        for key_node in resource_levels:
            node_id = ''
            # items w/same key_node will always have same value in 'Groups'
            group_list = self.spreadsheet_data['Resources'][sheet_tab_name][key_node][0]['Groups']
            parent_id = None
            for group in group_list:
                if len(node_id) == 0:
                    node_id = group
                else:
                    node_id += ',' + group
                this_node = self.resource_tree.get_node(node_id)
                if this_node is None:
                    try:
                        self.resource_tree.create_node(group,node_id,parent_id)
                    except MultipleRootError:  # MultipleRootError:
                        raise DXMLGeneratedError("[DXM-08] Group name '{0}' from xlsx in resource {1} not recognized".
                                                 format(node_id, sheet_tab_name))
                parent_id = node_id
        # self.resource_tree.show() - debug

    def _get_item_form_name(self, resource_name):
        """ Determine correct 'Form Name' in DDWikiImportConfig.xml to understand which fields are required in xlsx

        :param resource_name: Name of resource
        :return: (str) Value of attribute 'Name' in tag Form within DDWikiImportConfig.xml
        """
        try:
             full_resource_name = self.page_links[resource_name]
        except KeyError:
            raise DXMLGeneratedError("[DXM-47] Cannot find resource '{}' in config.ini [PageLinks] section".
                                     format(resource_name))
        if full_resource_name == 'Property Resource':
            return "PropResourceField"
        elif 'Collection' == full_resource_name.split()[-1]:
            return "CollectionResourceField"
        elif full_resource_name[0:6] == 'Lookup':
            return "LookupValue"
        else:
            return "OtherResourceField"

    def _create_resources(self):
        """ Build all XML nodes relating to Resources. Called when class is initialized

        :return (boolean): True/False on successful execution
        """
        # Resource sheets to grab from xlsx defined in config.ini
        for sheet_tab_name in self.program_config_data['ResourceSheets']:
            resource_name = self.program_config_data['ResourceSheets'][sheet_tab_name]
            self.logger.info("Processing Input Lookup Worksheet: '{}' for resource: '{}'".
                             format(sheet_tab_name, resource_name))
            # Build a tree structure for each resource which dups how the XML will be shaped
            self._build_resource_tree(sheet_tab_name=sheet_tab_name)

            self._create_resource_nodes(parent_xml_node=self.xml_root,
                                        sheet_tab_name=sheet_tab_name,
                                        resource_name=resource_name,
                                        group_name=resource_name,
                                        level_key=resource_name,
                                        config_form_name=self._get_item_form_name(resource_name))
        return True

    def _create_lookups(self):
        """ Build all Lookup XML nodes. Called when class is initialized

        :return: None
        """
        self.logger.info("Processing Input Lookup Values")
        # Create top node for Lookups
        top_lookup_node = self._add_xml_nodes(self.xml_root,
                                              nodes_from_config=self.xml_config_data["LookupTopIndex"],
                                              resource_name='Lookup')
        # Lookups grouped by 1st letter of lookup field
        for letter_key in sorted(self.spreadsheet_data['Lookups'].keys()):
            # Create Letter Group
            # Create alphabetic Lookup Indices
            page_title = self.xml_config_data["LookupIndexAlpha"]['Attributes']['Page_Title'].replace('[[Char]]',
                                                                                                      letter_key)
            top_group_index_node = self._add_xml_nodes(top_lookup_node,
                                                       nodes_from_config=self.xml_config_data["LookupIndexAlpha"],
                                                       other_page_title=page_title,
                                                       resource_name='Lookup Index')
            top_lookup_node.append(top_group_index_node)
            # Create a group node for each lookup field
            for lookup_field in sorted(self.spreadsheet_data['Lookups'][letter_key]):
                page_title = self.xml_config_data["LookupIndexField"]['Attributes']['Page_Title'].replace('[[Name]]',
                                                                                                lookup_field[0])
                labels=self.xml_config_data["LookupIndexField"]['Labels']['Value'].replace('[[alpha]]',
                                                                                           page_title[0].lower())
                # Add fields Translate <EnumerationID>> to LookupFieldID, <<lookupfield_ref>>
                # value picks up 1st lookup value item and picks off EnumerationID .. all have same value
                try:
                    val = lookup_field[1][0]['LookupFieldID']
                except KeyError:
                    val = None  # No NoLookupFieldID in .xlsx
                    # raise DXMLGeneratedError("[DXM-13] Error in accessing 'LookupFieldID' in Lookup tab")

                lookup_field_node = self._add_xml_nodes(top_lookup_node,
                                                        nodes_from_config=self.xml_config_data["LookupIndexField"],
                                                        value=val,
                                                        other_page_title=page_title,
                                                        replace_labels=labels,
                                                        resource_name='Lookup Field')
                top_group_index_node.append(lookup_field_node)
                # Add lookup Values
                for lookup_value in lookup_field[1]:
                    page_title = lookup_value['LookupValue']
                    lookup_value_node = self._add_xml_nodes(lookup_field_node,
                                                            nodes_from_config=self.xml_config_data["LookupValue"],
                                                            value=lookup_value,
                                                            other_page_title=page_title,
                                                            resource_name=lookup_field[0])

                    lookup_field_node.append(lookup_value_node)

    def write_xml_file(self, result_xml_filepath):
        """ Write IOI Import File to disk

        :param result_xml_filepath: (str) IOI Import filename/path
        :return: None. Raise DXMLGeneratedError on error
        """
        tree = xml_tree.ElementTree(self.xml_root)
        try:
            tree.write(result_xml_filepath, pretty_print=True)
        except (FileNotFoundError, IOError):
            raise DXMLGeneratedError("[DXM-11] Unable to write xml file: {}".format(result_xml_filepath))
        self.logger.debug("XML written to File:" + result_xml_filepath)

    def _read_max_ids(self, max_id_filepath):
        """ Parse through max id text file (stat_warning_log) and parse out max lookup id's

        :param max_id_filepath: (str) File/path of max id text file (stat_warning_log.txt)
        :return: None. Raise DXMLGeneratedError on error
        """
        self.max_lookupids = {}
        self.max_recordids = {}
        self.max_id = -1        # Look for a max record or lookup id
        max_id_file = ntpath.basename(max_id_filepath)
        try:
            gml_file = open(max_id_filepath, 'r')
        except FileNotFoundError:
            raise DXMLGeneratedError("[DXM-19] WikiStat File (Max Id's) not Found: " + max_id_filepath)
        gml_data = gml_file.readlines()
        gml_file.close()
        lookupid_section_found = False
        recordid_section_found = False
        for line in gml_data:
            if 'Max RecordID per Resource Report' in line:
                recordid_section_found = True
                lookupid_section_found = False
            elif 'Max LookupID per Lookup Field' in line:
                lookupid_section_found = True
                recordid_section_found = False
            else:
                if lookupid_section_found:
                    if line[0:2] != '**':
                        sline = line.split()
                        if len(sline) > 0:
                            if len(sline) != 5:
                                raise DXMLGeneratedError("[DXM-20] Max ID File has lookup text not recognized: " +
                                                         max_id_file)
                            val = int(sline[4])
                            try:
                                self.max_lookupids[sline[0]] = int(sline[4])
                            except ValueError:
                                raise DXMLGeneratedError("[DXM-21] Max ID File has illegal lookup id num for maxid: "
                                                         + sline[0])
                            if val > self.max_id:
                                self.max_id = val
                elif recordid_section_found:
                    if line[0:2] != '**':
                        sline = line.split()
                        if len(sline) > 0:
                            # Deprecated fields/lookups have more than 5 words
                            if sline[0] == 'Deprecated':
                                if sline[1] == 'Fields':
                                    del sline[1]
                                elif sline[1] == 'Lookup':
                                    sline[1:3] = []
                            if len(sline) != 5:
                                raise DXMLGeneratedError("[DXM-23] Max ID File expecting 5 cols per row for {} in:{}".
                                                         format(sline[0], max_id_file))
                            val = int(sline[4])
                            try:
                                self.max_recordids[sline[0]] = int(sline[4])
                            except ValueError:
                                raise DXMLGeneratedError("[DXM-24] Max ID File has illegal num for record maxid: "
                                                         + sline[0])
                            if val > self.max_id:
                                self.max_id = val

        if len(self.max_recordids) == 0:
            raise DXMLGeneratedError("[DXM-26] No recordid entries found in file " + max_id_file)
        if len(self.max_lookupids) == 0:
            raise DXMLGeneratedError("[DXM-09] No lookupid entries found in file " + max_id_file)
