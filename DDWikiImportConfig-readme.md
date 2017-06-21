# DDWikiImportConfig.xml Format
DDWikiImportConfig.xml dictates the output format for IOI Import xml file.
## Node Structure
* **DDWikiImportConfig.xml Purpose:** Describes content and format for output IOI Import xml file
* The XML Import File will have 2 types of XML nodes, 'Group' and 'Item'.
  * **Group**: A parent node corresponding to a summary Confluence page that is not a wiki field or lookup
  * **Item**: A detail node expressing either a wiki field or lookup value
* Each node in the XML file (*not including the root*) will correspond to a Wiki page and will appear in the DD Wiki navigation bar.
  * The IOI XML Import file is written the order as pages should appear in the final DD Wiki.
* XML Group Node
  * Groups are summary pages that contain Item  and other Groups nodes(e.g. children).
  * A Group will display as a parent tree node in the DD wiki navigation bar.  
* XML Item Node
  * XML Item nodes contain detailed information about a Data Dictionary field or lookup
  * Items Will Not contain other Group or Item nodes. 

### XML Construct
#### Parent Node - Form
The children nodes of 'Form' defines fields for a specific Confluence page type
* **Attribute Name:** Form Identifier
* **Attribute Page_Template:** Confluence Wiki Template Page Title
* **Attribute Page_Title:** DD Wiki Confluence Page Title 
* **Attribute Node_Type:** *Group* or *Item*

#### Child to Form - Field
* **Attribute Sequence:** Order this column will appear in output xml file
* **Attribute XMLName:** Name of column matched with Wiki Template for column placement
* **Attribute ParsingCode:** *(See Parse commands below)*
* **Attribute ChildTagName:** Name of child tag for complex node such as property types
* **Attribute AutoCompute:** Auto calculate field. Ignore any column in xlsx
* **Attribute DefaultValue:** Default value if xlsx column blank or missing
* **Value**: Name of corresponding column in xlsx sheet 

#### Child to Form - Labels
* **Attribute Sequence:** Order this column will appear in output xml file
* **Attribute XMLName:** Name of column matched with Wiki Template for column placement
* **Attribute ParsingCode:** *(See Parse commands below)*
* **Attribute ChildTagName:** Name of child tag for complex node such as property types

## Parse commands
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
    PARSE_LOOKUPID  = 10        # Compute LookupID
    PARSE_LOOKUP_FLDID = 11     # Compute LookupFieldID
    PARSE_RECORDID = 12         # Compute RecordID
    PARSE_FLD_REFERENCES = 13   # Parse Reference columns in Resource (collections)
    PARSE_FLD_COLLECTION = 14   # Parse a 'Collection' column to point to collection resource

Modification Date: *Jun 16 2017*