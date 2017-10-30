# IOI_Import.py
Converts MS Excel .xlsx files into .xml file (*IOI XML Import File*) for IOI to create confluence pages in temp work space.

## Input Spreadsheet (xlsx) files
* See gsheet template [DDWiki_1_7_Template v4 Sheets](https://drive.google.com/open?id=0B2lpfxXm4mUYMUVydkpVMkJqaWM)

## Program Command Line Arguments 
* -h, **--help**
* -f **--working_folder** <*current folder*>
  * Default folder for config and error log files                  
* -x **--xlsx_filename** <*none*>
  * Input .xlsx file containing new fields and lookups. File located in sub folder **StorageFiles** to working_folder.
  * Resultant **IOI xml import file** has same file name as xlsx file with xml used as extension
* -i **--max_id_filename** <*stat_warning_log.txt*>
  * Input file containing DD Wiki max record and lookup ids. Created by RESOExporter. 
  * File is in sub folder **StorageFiles** to working_folder.
* -w **--ddwiki_export_filename** <*none*>
  * Input file of latest DD Wiki exported XML file. Used to check for duplicate page titles.
  * Created by RESOExporter. 
  * File is in sub folder **StorageFiles** to working_folder.
* -d **--xlsx_date** <*today*>
  * Default date value (Status Change Date, Revised Date, Mod Date) for input spreadsheet (YYYY-MM-DD)
* -e **--error_logging** <*20*>
  * Error Logging Level (0-None, 10-Debug, 20-Info, 30-Warn, 40-Err, 50-Critical)

## Config.ini 
* **Purpose:** Describes how xlsx tabs/worksheets relate to DD Wiki resource/lookups, Resource/Collection definitions and page link text to specicif confluence pages.
* Section: **ResourceSheets**. key=.xlxs tab name, value=resource name in xml
  * Defines which resource/collection xlsx sheets (tabs) program should be transferred to IOI Import xml file
  * Notes on Resource fields with a Collection datatype
    * All fields within the same resource with a Collection datatype need to be grouped in same separate tab
    * Collections have a different wiki page format
    * A Collection sheet/tab name must end with ' Collection' otherwise assumed as Resource
    * The 'Collection' column in a xlsx Resource tab must be named 'Collection'
  * Collection are a type of Resource need to be grouped in separate tabs. The tab name must end with ' Collection'
* Section: **LookupSheets**. key=**LookupSheet**, value = .xlsx lookup tab name
  * Defines which Lookup sheet (tab) program should be transferred to IOI Import xml file
  * **note:** a xlsx can only have 1 sheet/tab dedicated to lookups
* Section: **Resource-Descriptions**. key=xml resource name, value=resource defintion as it will appear in wiki. The key should have the word ' Collection' for Collections otherwise a Resource is assumed. (' Resource' is not needed in key)
* Section: **PageLinks**. key=Display link text for Collection, Reference or Prop Type columns, value=Confluence Page Name to link
  * The Confluence Page Name has to match Confluence page title

## DDWikiImportConfig.xml and resultant (*e.g output*) xml file 
  * DDWikiImportConfig.xml describes format for resultant IOI_Import output xml file. See **DDWikiImportConfig-readme.md** for further detail.
  * The resultant IOI_Import XML are processing instructions on how to create Confluence pages
  * The resultant file will have 2 types of XML tags, 'Group' and 'Item'.
    * **Group**: A parent node corresponding to a summary Confluence page that is not a field or lookup
      * Groups are summary pages that represent collections (e.g. children) of Items and other Groups.
      * A Group will display as a tree node in the DD wiki navigation bar.  
    * **Item**: A detail node expressing either a wiki field or lookup value
      * XML Item nodes contain detailed information about a Data Dictionary field or lookup
      * Items Will Not contain other Group or Item nodes.
  * Each node in the resultant XML file (*not including the root*) will correspond to a Wiki page and will appear as tree nodes in the DD Wiki navigation bar.
	* The IOI XML Import file is already structured and in the order as it should appear in the final DD Wiki pages
  * Confluence Notes for programming import process: [DD Wiki IOI Import Process - Phase 3b](https://goo.gl/XHTlmN)

## Other Notes of Importance
### Copy latest exported xml and wiki stat file
* IOI_Import requires two files created by the RESOConfluenceExporter project. They are:
  * Exported XML File: DD Wiki Representation in xml
  * DD Wiki Stat File: Text file listing max id numbers used in Resources, Collections and Lookup Fields
  
Last Modification Date: *Oct 26 2017*  



