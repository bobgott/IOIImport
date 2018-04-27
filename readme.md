# Creating an IOI Import XML File
This project converts RESO Data Dictionary New Fields/Lookups from an .xlsx file into IOI .xml import file for IOI to create confluence pages in a temporary Confluence work space.

## Input Spreadsheet (xlsx) files
Workbook (.xlsx) must be in the format as dictacted in the template below. Instructions are also included in the template workbook. All rows in a worksheet tab must belong to the same Resource
* See template [DDWiki_1_7_Template v8 Sheets](https://drive.google.com/file/d/1h8LbdsWnbh1To1IGRJOs3T-6RxHcToor/view?usp=sharing) dated Apr 26 2018.

## Program Command Line Arguments 
* -h, **--help**
* -f, **--home_folder** <*current folder*>
  * Default root folder for all applications, configuration files, error log files, etc.
* -c, **--config_sub_folder** <'current'>
  * Sub Folder under 'files' then 'config' containg configuration and ini files                  
* -x, **--xlsx_filename** <*none*>
  * Input .xlsx file containing new fields and lookups to be imported into DD Wiki. 
  * File located under 'files' then 'input' folder.*
  * **Note:** Resultant/Output file for IOI has same file name as xlsx file but using .xml as file extension and located under 'files' then 'xml' folder.
* -i, **--max_id_filename** <*stat_warning_log.txt*>
  * Input file containing DD Wiki max record and lookup ids. File created by WikiExporter. 
  * File located under 'files' then 'input' folder. *
* -w, **--ddwiki_exported_xml_filename** <*none*>
  * Input xml file from created by WikiExported. Used to check for duplicate page titles. File created by WikiExporter. 
  * File located under 'files' then 'input' folder. *
* -d, **--xlsx_date** <*today*>
  * Default date value for *Status Change Date, Revised Date, Mod Date* for Resultant/Output file
* -e, **--error_logging** <*20*>
  * Error Logging Level (0-None, 10-Debug, 20-Info, 30-Warn, 40-Err, 50-Critical)

## Config.ini 
* **Purpose:** Describes how input xlsx tabs/worksheets relate to the resultant DD Wiki IOI import xml file. 
* Section: **ResourceSheets** - key: .xlxs sheet/tab name, value: DD Wiki Resource name
  * Maps xlsx sheets (tabs) to DD Wiki resource or collections that will be transferred to resultant IOI Import xml file and hence the DD Wiki confluence pages.
  * Notes on Resource fields with a Collection datatype
    * All fields within the same resource or collection need to be grouped in same tab.
* Section: **LookupSheets** - key:**LookupSheet**, value: .xlsx lookup tab name *(note: different order than [ResourceSheets])*
  * Defines which sheet (tab) program should be transferred to Lookup Fields and Values for the IOI Import xml file
  * **note:** A xlsx file can only have 1 sheet/tab dedicated to lookups. 
* Section: **Resource-Descriptions** - key=xml resource/collection name, value=resource/collection defintion as it will appear in wiki.
* Section: **PageLinks** - key: xml resource/collection name, value: Confluence Page Name
  * The Confluence Page Name has to match Confluence page title. This also identifies between resource or collection

## DDWikiImportConfig.xml and resultant (*e.g output*) IOI import xml file 
  * DDWikiImportConfig.xml describes format for resultant IOI import output xml file. See **DDWikiImportConfig-readme.md** for further detail.
  * The resultant IOI import xml file contains processing instructions on how IOI should create DD Wiki Confluence pages
  * The resultant file will have 2 types of XML tags, 'Group' and 'Item'.
    * **Group**: A parent node corresponding to a summary Confluence page that is not a field or lookup
      * Groups are summary pages that represent collections (e.g. children) of Items and other Groups.
      * A Group will display as a tree node in the DD wiki navigation bar.  
    * **Item**: A detail node expressing either a wiki field or lookup value
      * XML Item nodes contain detailed information about a Data Dictionary field or lookup
      * Items Will Not contain other Group or Item nodes.
  * Each node in the resultant XML file (*not including the root*) will correspond to a Wiki page and will appear as tree nodes in the DD Wiki navigation bar.
	* The IOI XML Import file is structured and in the order as it should appear in the final DD Wiki pages

## Other Notes of Importance
### Prior to running progra, copy latest exported xml and wiki stat file
* IOI_Import requires two files created by the WikiExporter project. They are:
  * Exported XML File: DD Wiki Representation in xml
  * DD Wiki Stat File: Text file listing max id numbers used in Resources, Collections and Lookup Fields
  
Last Modification Date: *Apr 26 2018*  



