

Public Function Backup_BE_Tables(BE_Directory As String, Optional ByVal SaveLocation As String, Optional ByVal Save_Identifier As String) As String
'Backup all tables in database to specified folder. 
'Current functions uses a Load_Globals function that you will need to create or deploy another method to gather current DB path like currentdb. etc
On Error GoTo ErrorHandler
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    Dim CurrentDateTime As String
    Dim BackupFolderPath As String
    Dim BackupLog As String
    Dim Current_BE_Filepath As String
    Dim Current_BE_path As String
    Dim Current_BE_FileName As String
    
    Call Load_Globals 'I always store Backend information in global variables 'Filepath', filename, etc. I'm ensuring they are loaded here. You may choose to do it diff)
    
    CurrentDateTime = format(Now(), "yyyymmdd_hhmmss")
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb
    For Each tdf In db.TableDefs
        DoCmd.Hourglass True
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then   ' Or tdf.NAME = "Temp_Tables.accdb"
            Current_BE_Filepath = Mid(CurrentDb.TableDefs(tdf.Name).Connect, InStrRev(CurrentDb.TableDefs(tdf.Name).Connect, "=") + 1)
            If (InStr(CurrentDb.TableDefs(tdf.Name).Connect, "WSS;HDR=NO;IMEX=2;ACCDB=YES;DATABASE=") = 0) And (InStr(BackupLog, Current_BE_Filepath) = 0) Then
                If InStr(Current_BE_Filepath, "Temp_Tables.accdb") = 0 And Not isBlankOrNull(Current_BE_Filepath) Then
                    Current_BE_FileName = Replace(Current_BE_Filepath, BE_Directory, "")  'skip backup of temp tables backend
                    If isBlankOrNull(SaveLocation) Then
                        BackupFolderPath = BE_Directory & "Backups"
                        Call CreateDirectory(BackupFolderPath)
                        BackupFolderPath = BackupFolderPath & "\" & CurrentDateTime
                        If Not isBlankOrNull(Save_Identifier) Then BackupFolderPath = BackupFolderPath & " - " & Save_Identifier
                        Call CreateDirectory(BackupFolderPath)
                    Else
                        BackupFolderPath = SaveLocation
                        Call CreateDirectory(BackupFolderPath)
                    End If
                    Call CopyFile(Current_BE_Filepath, BackupFolderPath, "\" & Current_BE_FileName)
                    BackupLog = BackupLog & Current_BE_Filepath
                End If
            End If
        End If
    Next

Backup_BE_Tables = CurrentDateTime
    
ErrorHandler:
    If (Err.Number > 0) Then
        Backup_BE_Tables = -1
        Select Case Err.Number
            Case Else
                Call ErrorHandelerMsgBoxLog(Err.Number, Err.Description, "Backup_Local_BE_Tables")
        End Select
        Resume Next
    End If
CleanExitTask:

Set tdf = Nothing
Set db = Nothing
DoCmd.SetWarnings True
DoCmd.Hourglass False
End Function



Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
'Changes access option properties. Usefull when automating the creation of an ACCDE file
    Dim dbs As Object, prp As Variant
    Const conPropNotFoundError = 3270
    Set dbs = CurrentDb
    On Error GoTo Change_Err
    dbs.Properties(strPropName) = varPropValue
    ChangeProperty = True
Change_Bye:
    Exit Function
Change_Err:
    If Err = conPropNotFoundError Then ' Property not found.
        Set prp = dbs.CreateProperty(strPropName, _
            varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else
        ' Unknown error.
        ChangeProperty = False
        Resume Change_Bye
    End If
End Function



Public Function CloseAllOpenForms()
'Close all open Forms
'You can exclude forms by adding them to the if statment below
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject
For Each obj In dbs.AllForms
  If obj.IsLoaded = True Then
    If (obj.Name <> "ShutdownMonitor" And obj.Name <> "NotificationWindow") Then 'omit form names here
        DoCmd.Close acForm, obj.Name, acSaveYes
    End If
  End If
Next obj
End Function



Public Function CompactAndRepairAll(BE_Directory As String, BE_Backup_Directory As String) As Boolean
'Automates the compact and repair of all linked tables in a database. The user must still click an 'ok' button for every linked Backend file associated with tables.
'You must verify that the database is not in use by other users before calling.
'Current functions uses a Load_Globals function that you will need to create or deploy another method to gather current DB path like currentdb. etc
'Dependent on RepairDatabase Function
'    Call Load_Globals
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    Dim CurrentDateTime As String
    Dim BackupFolderPath As String
    Dim CompactLog As String
    Dim FailFlag As Boolean
    Dim Current_BE_Filepath As String
    Dim Current_BE_path As String
    Dim Current_BE_FileName As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb
	CurrentDateTime = format(Now(), "yyyymmdd_hhmmss")
    For Each tdf In db.TableDefs
        DoCmd.Hourglass True
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
            Current_BE_Filepath = Mid(CurrentDb.TableDefs(tdf.Name).Connect, InStrRev(CurrentDb.TableDefs(tdf.Name).Connect, "=") + 1)
            If (InStr(CurrentDb.TableDefs(tdf.Name).Connect, "WSS;HDR=NO;IMEX=2;ACCDB=YES;DATABASE=") = 0) And (InStr(CompactLog, Current_BE_Filepath) = 0) And (Not isBlankOrNull(Current_BE_Filepath)) Then
                Current_BE_FileName = Replace(Current_BE_Filepath, BE_Dir, "")
                If Not (Current_BE_FileName Like "Temp_Tables.accdb") Then
                    FailFlag = FailFlag + Not RepairDatabase(BE_Directory & Current_BE_FileName, BE_Backup_Directory & Current_BE_FileName)
                    CompactLog = CompactLog & Current_BE_Filepath
                End If
            End If
        End If
    Next
    CompactAndRepairAll = Not FailFlag
CleanExitTask:
Set tdf = Nothing
Set db = Nothing
DoCmd.SetWarnings True
DoCmd.Hourglass False
End Function



Public Function Create_ACCDE_File(Optional ByVal ConvertToACCDR As Boolean = False, Optional RemoveDebug as boolean = False) As Boolean
'automatically creates an ACCDE File with appropriate settings and such for a clean accde view
    Dim app As New Access.Application
    Dim ACCDE_FilePath As String
    Dim ACCDE_FileName As String

    if RemoveDebug then RemoveDebugs		'VBA IDE Foundation. removes debug rows from code in IDE.
    
    ACCDE_FileName = Left(CurrentProject.Name, Len(CurrentProject.Name) - 6)
    ACCDE_FilePath = CurrentProject.Path & "\For Release\"
    
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accdb") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accdb"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE") Then Kill ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accde") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accde"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accdr") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accdr"
    
    Call CopyFile(CurrentProject.FullName, ACCDE_FilePath, ACCDE_FileName & ".accdb")

    app.AutomationSecurity = 1 ' this is msoAutomationSecurityLow -- can't figure out what the library constant is, so I just used 1.
    app.SysCmd 603, ACCDE_FilePath & ACCDE_FileName & ".accdb", ACCDE_FilePath & ACCDE_FileName & ".accde"
    
    If Not FileExists(ACCDE_FilePath & ACCDE_FileName & ".accde") Then
        Create_ACCDE_File = False
    Else
        Name ACCDE_FilePath & ACCDE_FileName & ".accdb" As ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE"
        If ConvertToACCDR Then Name ACCDE_FilePath & ACCDE_FileName & ".accde" As ACCDE_FilePath & ACCDE_FileName & ".accdr"
        Create_ACCDE_File = True
    End If
End Function



Public Function CreateTableArray(ByRef SQL_String As String) As Variant
'Conversts a table into an array.
Dim rstData As DAO.Recordset
If (Not isBlankOrNull(SQL_String)) Then
'    'debug.print SQL_String
    Set rstData = CurrentDb.OpenRecordset(SQL_String)
    If rstData.BOF = False Then
        rstData.MoveLast
        rstData.MoveFirst
        CreateTableArray = rstData.GetRows(rstData.RecordCount) 'Stores in multidimensional array reguardless of column count!!!!!!!!!!!!!
        rstData.Close
        If IsEmpty(CreateTableArray) Then
            CreateTableArray = -1
        Else
            CreateTableArray = TransposeArray(CreateTableArray)
        End If
    Else
        CreateTableArray = -2
    End If
Else
    CreateTableArray = -3
End If
Set rstData = Nothing
End Function



Public Function ErrorHandelerMsgBoxLog(ErrNumber As Long, ErrDesc As String, Optional SubroutineName As String = "Unknown", Optional LogToTxtFile As Boolean = True, Optional LogToErrorTable As Boolean = False)
'Use this in your error handle routines instead of msgbox. Write errors to txt file and or a Table. Passes error to msgbox for user visibility.
Dim db As DAO.Database
Set db = CurrentDb
Dim SQL_String As String

If ErrNumber = 3343 And InStr(ErrDesc, "CID v") > 0 Then
   MsgBox "Access file has become corrupt. Contact Support/Supervisor"
   CloseDatabase
Else
   MsgBox "Error# " & ErrNumber & " : " & ErrDesc & vbNewLine & vbNewLine & _
   "Subroutine: " & SubroutineName & vbNewLine & vbNewLine & _
   "Saved to Log File."
End If

ErrDesc = Replace(ErrDesc, vbNewLine, " ")
ErrDesc = RemoveExtraSpaces(ErrDesc)
ErrDesc = Replace(ErrDesc, "'", "|")
ErrDesc = Replace(ErrDesc, """", "|")

If LogToErrorTable = True Then
SQL_String = "INSERT INTO ErrorLogs ( DateTimeStamp, SubroutineName, ErrNumber, ErrDescription ) SELECT '" & Now() & "' AS Expr1, '" & SubroutineName & "' AS Expr2, '" & ErrNumber & "' AS Expr3, '" & ErrDesc & "' AS Expr4;"
db.Execute SQL_String
End If

If LogToTxtFile Then Call AppendTXTFile(Now() & " - Subroutine: " & SubroutineName & "|Error# " & ErrNumber & " : " & ErrDesc, BE_Dir & "logfile.txt")
End Function




Public Function ExportTableExcel(TableName As String, ExportFileNamePath As String, Optional ColumnExcludes As String, Optional ExportHeaderNames As Boolean = True, Optional ExportFormat As AcSpreadSheetType = acSpreadsheetTypeExcel9, Optional WorksheetName As String)
'----ColumnExcludes must be comma seperated ("Column_2,Column_5")
Dim db As DAO.Database
Dim SQL_String As String
Dim rst As New ADODB.Recordset
Dim ColumnList As String
Dim i As Integer

Set db = CurrentDb

If Not isBlankOrNull(ColumnExcludes) Then

    SQL_String = "select * from  " + TableName + " where 1=0"
    rst.Open SQL_String, CurrentProject.Connection
    ColumnList = "\*"
    For i = 0 To rst.Fields.Count - 1
        If InStr(ColumnExcludes, rst.Fields(i).Name) = 0 Then
            ColumnList = ColumnList & "," & rst.Fields(i).Name
        End If
    Next i
    ColumnList = Replace(ColumnList, "\*,", "")
    rst.Close
    SQL_String = "Select " & ColumnList & " From " & TableName
    Set Qdf = db.CreateQueryDef("Export_Temp_Table", SQL_String)
    DoCmd.TransferSpreadsheet acExport, ExportFormat, "Export_Temp_Table", ExportFileNamePath, ExportHeaderNames, WorksheetName
Else
    DoCmd.TransferSpreadsheet acExport, ExportFormat, TableName, ExportFileNamePath, ExportHeaderNames, WorksheetName
End If




Public Function IsFormOpen(FormName) As Boolean
'Checks is form is already open
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject
IsFormOpen = False
For Each obj In dbs.AllForms
  If obj.IsLoaded = True Then
    If (obj.Name = FormName) Then
        IsFormOpen = True
    End If
  End If
Next obj
End Function



Public Function IsFormLoaded(strFormName As String, Optional LookupFormName As Form, Optional LookupControl As Control) As Boolean
'A more robust version of isformopen, but allows the detection of subforms. Use LookupFormName and LookupControl to speed up subform open find.
    Dim frm As Form
    Dim bFound As Boolean
 
    If Not (IsMissing(LookupFormName) Or IsNull(LookupFormName) Or LookupFormName Is Nothing) Then
        If Not (IsMissing(LookupControl) Or IsNull(LookupControl) Or LookupControl Is Nothing) Then
        On Error GoTo ErrorHandler:
            If LookupControl.ControlType = acSubform Then
                If LookupControl.Form.Name = strFormName Then
                    bFound = True
                End If
            End If
        Else
            Call SearchInForm(LookupFormName, strFormName, bFound)
        End If
    Else
        For Each frm In Forms
            If frm.Name = strFormName Then
                bFound = True
                Exit For
            Else
                Call SearchInForm(frm, strFormName, bFound)
                If bFound Then
                    Exit For
                End If
            End If
        Next
    End If
    IsFormLoaded = bFound
End Function


Public Function ListboxSelectionClear(ByRef ListName As Control)
'Clears Listbox selections
    Dim varItm As Variant
    With ListName
        For Each varItm In .ItemsSelected
            .Selected(varItm) = False
        Next varItm
    End With
End Function


Public Function ListboxSelectionArray(ByRef ListName As Control) As Variant
'Transfer selected items in a listbox to an array
Dim varItm As Variant, intI As Integer
Dim totalSelected As Long
Dim Totalcolumns As Integer
Dim tempArray As Variant
Dim i As Long

totalSelected = ListName.ItemsSelected.Count
Totalcolumns = ListName.ColumnCount

If totalSelected > 0 And Totalcolumns > 0 Then
    ReDim tempArray(1 To totalSelected, 1 To Totalcolumns)
    For Each varItm In ListName.ItemsSelected
        i = i + 1
        For intI = 0 To ListName.ColumnCount - 1
            tempArray(i, intI + 1) = ListName.column(intI, varItm)
        Next intI
    Next varItm
End If
If Not isBlankOrNull(tempArray) Then ListboxSelectionArray = tempArray
End Function


Public Function RelinkBackend(LinkedTableName As String, Path As String, BackendFileNamePath As String) As Boolean
'Relink the tables to new backend filenamepath
Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
Dim strTable As String
Set dbs = CurrentDb()
For Each tdf In dbs.TableDefs
    If (tdf.Name = LinkedTableName) Then
        If tdf.Connect <> ";DATABASE=" & Path & BackendFileNamePath Then
            dbs.TableDefs(tdf.Name).Connect = ";DATABASE=" & Path & BackendFileNamePath
            dbs.TableDefs(tdf.Name).RefreshLink
        End If
    End If
Next tdf
RelinkBackend = True
End Function




Public Sub RefreshODBCLinks(newConnectionString As String)
'Refresh ODBC Connections (this may not do what you are trying to do. It is usefule in certain instances like refreshing to update new or deleted tables)
    Dim db As DAO.Database
    Dim tb As DAO.TableDef
    Set db = CurrentDb
    For Each tb In db.TableDefs
        Debug.Print tb.Connect
        If Left(tb.Connect, 4) = "ODBC" Then
            Debug.Print tb.Name
			tb.Connect = newConnectionString
			tb.RefreshLink
        End If
    Next tb
    Set db = Nothing
End Sub




Function RepairDatabase(strSource As String, strDestination As String) As Boolean.
'Automates the compact and repair of specified linked table. The user must still click an 'ok' button.
    On Error GoTo Error_Handler
    RepairDatabase = _
        Application.CompactRepair( _
        LogFile:=True, _
        SourceFile:=strSource, _
        DestinationFile:=strDestination)
    On Error GoTo 0
    Exit Function
Error_Handler:
    RepairDatabase = False
End Function



Public Function VerifyReqFields(ReqFieldsList As Variant) As Boolean
'Validates all required fields on a form are not blank. If any are blank tells use which field it is returns false.
For Each item In ReqFieldsList
    If isBlankOrNull(item.Value) Then
        VerifyReqFields = False
        Call MsgBox(item.Name & " - is a required field." & vbNewLine & vbNewLine & "Correct missing information and try again.", , "Required Fields Missing")
        Exit For
    Else
        VerifyReqFields = True
    End If
Next item
End Function
