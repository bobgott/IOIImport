# DDWikiImportConfig.xml Format
DDWikiImportConfig.xml describes content and format for IOI_Import output xml file.
## *Background:* Node Structure for the IOI_Import output xml file
* The IOI_Import output xml file will have 2 types of XML nodes, 'Group' and 'Item'.
    * **Group**: A parent node corresponding to a summary Confluence page that is not a field or lookup
      * Groups are summary pages that represent collections (e.g. children) of Items and other Groups.
      * A Group will display as a tree node in the DD wiki navigation bar.  
    * **Item**: A detail node expressing either a wiki field or lookup value
      * XML Item nodes contain detailed information about a Data Dictionary field or lookup
      * Items Will Not contain other Group or Item nodes.
* Each node in the XML file (*not including the root*) will correspond to a Wiki page and will appear in the DD Wiki navigation bar.
  * The IOI XML Import file is written in the order as pages should appear in the final DD Wiki.

### DDWikiImportConfig.xml Construct
#### **Tag: Form** - Defines fields for a Confluence page type
* **Attribute Name:** Form Identifier
* **Attribute Page_Template:** Confluence Wiki Template Page Title
* **Attribute Page_Title:** DD Wiki Confluence Page Title 
  * [[*xxx*]] - Replace [[*xxx*]] with name of Resource or Collection
* **Attribute Node_Type:** *Group* or *Item*

#### **Tag: Field** - Describes all fields within a Form
* **Attribute Sequence:** Order this column will appear in output xml file (*order is not important, but helpful in debugging*)
* **Attribute XMLName:** Name of column matched with Wiki Template for column
* **Attribute ParsingCode:** *(See Parse commands below)*
* **Attribute ChildTagName:** Name of child tag for complex node such as property types
* **Attribute AutoCompute:** Auto calculate field. Ignore any column in xlsx
* **Attribute CollectionTemplate:** If a field references a Collection Resource, then a different template is used
* **Attribute DefaultValue:** Default value if xlsx column blank or missing
* **Value**: Name of corresponding column in xlsx sheet 

#### **Tag: Labels** - Describes Confluence labels to be applied to Form
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

Modification Date: *Oct 26 2017*