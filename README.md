# VBA_Foundation_Functions
A Repo of VBA functions I have written over the years that provide a powerful base for VBA Development.

I like to think of all the functions in this libary as foundation functions that should have been built into VB in the first place. It's broken into the following catagories:

1. General: VBA Functions that could probably be used in any app.
2. Access: Specific To Access.
3. Access_Excel_Interact: Used for Excel Doc interaction from within Access.

*Notes:*
1. *I intentionally do not use OnError unless its required to make the function operate. I perfer to handle errors in the main function that called these foundation functions. Just pretend these are built into VBA and handle errors like you normally would.*
2. *I intentionally leave all variable declarations on a seperate lines. These are functions I have written or collected over a long period of time. I find it helps to break everything down just in case i need to look into how something works again.*

**General Functions**
- AppendTXTFile: Add passed string to end of txt file.
- ArraySearch: Find a value in a specific Array column.
- ArraySearchComp: Find a value in a specific Array column with textcompare.
- ArraySearch_All: Checks if value exist in any Array column.
- BrowseForFile: Create a browse for file dialog box. returns selected filename.
- BrowseForFolder: Create a browse for folder dialog box. returns selected folder.
- CreateDirectory: Create Dir if not exists.
- CopyFile: Creates a copy of a file.
- CreateFolder: if folder does not exist create
- DeleteFile: Deletes File if exists.
- DeleteFolder: Deletes Folder if exists.
- DebugPrintArray: Debug print entire array.
- ForceCompileProject: Complies your code programmatically.
- FileExists: Check if file exist.
- GetBetween: Returns strings inbetween two strings.
- GetFileCount: Returns total files in folder.
- GetFilenameFromPath: Returns the rightmost characters of a string upto but not including the rightmost '\'.
- IsBlankOrNull: Returns true if passed value is null or "" reguardless of variable type. I use this so much I almost forget its not a built in VBA Function.
- IsFileOpen: Verifies if file is actively open on this or another computer.
- IsNumberKeyPress: Verifies if number key on number key row or numberpad is pushed. Usefill in certain situations to only allow number key push on number only fields.
- IsProcessRunning: Use to check active windows processes.
- openfile: Open file if exists.
- Pause: Waits X time in seconds.
- RandomNumberBetween: Returns random number between uper and lower passed values.
- RandomWeightedNumber: Returns random weighted value from array.
- RemoveExtraSpaces: Removes all double spaces from passed string.
- TransposeArray: Swaps column and rows of an array.
- WaitForTime: Waits until the specified date and time.

**Access Functions**
- Backup_BE_Tables: Backup all tables in database to specified folder. 
- ChangeProperty: Changes access option properties. Usefull when automating the creation of an ACCDE file.
- CloseAllOpenForms: Close all open Forms.
- CompactAndRepairAll: Automates the compact and repair of all linked tables in a database. The user must still click an 'ok' button for every linked Backend file associated with tables.
- Create_ACCDE_File: Automatically creates an ACCDE File with appropriate settings and such for a clean accde view.
- CreateTableArray: Conversts a table into an array.
- ErrorHandelerMsgBoxLog: Use this in your error handle routines instead of msgbox. Write errors to txt file and or a Table. Passes error to msgbox for user visibility.
- ExportTableExcel: Exports table to Excel File with option to exclude columns by name.
- IsFormOpen: Checks is form is already open.
- IsFormLoaded: A more robust version of isformopen, but allows the detection of subforms. Use LookupFormName and LookupControl to speed up subform open find.
- ListboxSelectionClear: Clears Listbox selections.
- ListboxSelectionArray: Transfer selected items in a listbox to an array.
- RelinkBackend: Relink the tables to new backend filenamepath.
- RefreshODBCLinks: Refresh ODBC Connections.
- RepairDatabase: Automates the compact and repair of specified linked table. The user must still click an 'ok' button.
- VerifyReqFields: Validates all required fields on a form are not blank. If any are blank tells user which field it is returns false.

**Access_Excel_Interact**
- CloseAllExcel: Close all open Excel workbooks.
- CloseExcelFile: Close specific Excel workbook based on filenamepath.
- IsExcelOpen: Returns true if an instance of Excel is Open.
- IsWorkbookOpen: Returns true if specified workbook is open.
- IsUserCellChangesUncommited: Determines if user has entered info into a cell, but has not yet commited the data. Returns True if it is uncommited.
- OpenWorkbook: Open an Excel Workbook based on filenamepath.


