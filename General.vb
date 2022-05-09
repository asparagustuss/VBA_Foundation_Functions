Option Explicit

Public Const vbDoubleQuote As String = """"

Public Function AppendTXTFile(LineToBeWritten As String, sFilePath As String, Optional ByVal WriteData As Boolean)
'add passed string to end of txt file.
'Default Print (for literal string writes) - https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/printstatement
'Optional Write (for data writes) - https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/writestatement
Dim FileNumber As Long
    FileNumber = FreeFile
    If (Len(Dir(sFilePath))) = 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim oFile As Object
        Set oFile = fso.CreateTextFile(sFilePath)
        Set fso = Nothing
        Set oFile = Nothing
    End If
    Open sFilePath For Append As #FileNumber
    If WriteData = True Then
        Write #FileNumber, LineToBeWritten
    Else
        Print #FileNumber, LineToBeWritten
    End If
    Close #FileNumber
End Function

Public Function RemoveExtraSpaces(CleanupString As String) As String
'removes all double spaces from passed string
Do While InStr(CleanupString, "  ") <> 0
    CleanupString = Replace(CleanupString, "  ", " ")
Loop
RemoveExtraSpaces = CleanupString
End Function



Public Function ArraySearch(arr As Variant, sFindValue As Variant, column As Long) As Long
'find a value in a specific Array column
Dim i As Long
TwoD_ArraySearch = -1

For i = LBound(arr) To UBound(arr)
    If arr(i, column) = sFindValue Then
        TwoD_ArraySearch = i
        Exit For
    End If
Next i
End Function



Public Function ArraySearchComp(arr As Variant, sFindValue As String) As Long
'find a value in a specific Array column with textcompare
Dim i As Long
ArraySearch = -1

For i = LBound(arr) To UBound(arr)
    If StrComp(sFindValue, arr(i), vbTextCompare) = 0 Then
        ArraySearch = i
        Exit For
    End If
Next i
End Function



Public Function ArraySearch_All(ByVal arValues As Variant, ByVal sFindValue As Variant) As Long
'checks if value exist in any Array column
Dim lrow As Long
Dim lcolumn As Long
TwoD_ArraySearch_All = True
    For lrow = LBound(arValues, 1) To UBound(arValues, 1)
        For lcolumn = LBound(arValues, 2) To UBound(arValues, 2)
            If (arValues(lrow, lcolumn) = sFindValue) Then
                TwoD_ArraySearch_All = True
                Exit Function
            End If
        Next lcolumn
    Next lrow
End Function


Public Function BrowseForFile(Optional TitleName As String = "Select File", Optional Button_Name As String = "Select", Optional File_Filter As String = "", Optional MultiSelect As Boolean = False) As Variant
'create a browse for file dialog box. returns selected filename.
	Dim objFileDialog As Object
    Dim i As Long
    Dim varItem As Variant
    Set objFileDialog = Application.FileDialog(3)
    With objFileDialog
        .ButtonName = "Select"
        .AllowMultiSelect = MultiSelect
        .Filters.Clear
        If Not isBlankOrNull(File_Filter) Then
            .Filters.Add "Limited To", File_Filter  '<--- must contain * ("*.txt")
        End If
        .Title = TitleName
        .Show
        If .SelectedItems.Count > 1 Then
            ReDim SelectedFiles(1 To .SelectedItems.Count) As Variant
            For Each varItem In .SelectedItems
                i = i + 1
                SelectedFiles(i) = varItem
            Next varItem
            BrowseForFile = SelectedFiles
        ElseIf (.SelectedItems.Count = 1) Then
            BrowseForFile = .SelectedItems(.SelectedItems.Count)
        Else
            BrowseForFile = ""
        End If
    End With
End Function


Public Function BrowseForFolder(Optional OpenAt As Variant, Optional Options As Long, Optional Title As String) As Variant
'create a browse for folder dialog box. returns selected folder.
    'Options: hex numbers on page must be convereted into decimals and added together and passed under Options argument
    'https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/ns-shlobj_core-browseinfoa
    Dim ShellApp As Object
    If isBlankOrNull(Title) Then Title = "Please choose a folder"
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, Title, Options, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":"
            If Left(BrowseForFolder, 1) = ":" Then BrowseForFolder = False
        Case Is = "\"
            If Not Left(BrowseForFolder, 1) = "\" Then BrowseForFolder = False
        Case Else
            BrowseForFolder = False
    End Select
End Function


Public Function CreateDirectory(DirPath As String)
'Create Dir if not exists
    If Dir(DirPath, vbDirectory) = "" Then
        MkDir DirPath
    End If
End Function



Public Function CopyFile(ByVal Origin_FileNamePath As String, ByVal Destination_filePath As String, ByVal Destination_FileName As String) As Variant
'creates a copy of a file
Dim fso As Object
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Destination_filePath) Then fso.CreateFolder (Destination_filePath)
CopyFile = fso.CopyFile(Origin_FileNamePath, Destination_filePath & Destination_FileName, 1)
End Function


Public Function DeleteFile(ByVal FileToDelete As String)
'Deletes File if exists
   If FileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Function



Public Function DeleteFolder(ByVal FolderToDelete As String)
'Deletes Folder if exists
Dim fso As Object
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
If FileExists(FolderToDelete) Then
    fso.DeleteFolder (FolderToDelete)
End If
Set fso = Nothing
End Function


Public Function DebugPrintArray(arr As Variant)
'use to debug print an array
Dim lrow As Long
Dim lcolumn As Long
For lrow = LBound(arr, 1) To UBound(arr, 1)
    For lcolumn = LBound(arr, 2) To UBound(arr, 2)
        Debug.Print arr(lrow, lcolumn)
    Next lcolumn
Next lrow
End Function

Public Function Force_CompileProject() As Boolean
'complies your code programmatically
DoCmd.RunCommand acCmdCompileAndSaveAllModules
Force_CompileProject = Application.IsCompiled
End Function



Public Function FileExists(ByVal FileToTest As String) As Boolean
'Check if file exist
On Error GoTo ErrorHandler
   FileExists = (Dir(FileToTest, vbDirectory) <> "")
ErrorHandler:
If (Err.Number > 0) Then
    Select Case Err.Number
        Case 52
            FileExists = False
    End Select
    Resume Next
End If
End Function


Public Function FolderExistsCreate(DirectoryPath As String, CreateIfNot As Boolean) As Boolean
'if folder does not exist create
    Dim Exists As Boolean
    On Error GoTo DoesNotExist
    Exists = ((GetAttr(DirectoryPath) And vbDirectory) = vbDirectory)

    If Exists Then
        FolderExistsCreate = True
    Else
        If CreateIfNot Then
            MkDir DirectoryPath
            FolderExistsCreate = True
        Else
            FolderExistsCreate = False
        End If
    End If
    Exit Function
    
DoesNotExist:
    FolderExistsCreate = False
End Function



Public Function Force_CompileProject() As Boolean
'complies your code programmatically
DoCmd.RunCommand acCmdCompileAndSaveAllModules
Force_CompileProject = Application.IsCompiled
End Function



Public Function GetBetween(ByRef sSearch As String, ByRef sStart As String, ByRef sStop As String, Optional ByRef lSearch As Long = 1) As String
'returns strings inbetween two strings
    lSearch = InStr(lSearch, sSearch, sStart)
    If lSearch > 0 Then
        lSearch = lSearch + Len(sStart)
        Dim lTemp As Long
        lTemp = InStr(lSearch, sSearch, sStop)
        If lTemp > lSearch Then
            GetBetween = Mid$(sSearch, lSearch, lTemp - lSearch)
        End If
    End If
End Function




Public Function GetFileCount(folderspec As String) As Integer
' better than using DoEvents which eats up all the CPU cycles
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.FolderExists(folderspec) Then
      GetFileCount = fso.GetFolder(folderspec).Files.Count
   Else
      GetFileCount = -1
   End If
End Function


Public Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function



Public Function GetRandomWeightedNo1(ByVal TotalChoices As Integer, ByVal ChoiceWeights As Variant) As Integer
'returns random weighted value
Dim SumOfWeight As Integer
Dim RandomSelection As Integer
Dim i As Integer

SumOfWeight = 0
For i = 0 To TotalChoices
   SumOfWeight = SumOfWeight + ChoiceWeights(i)
Next i
RandomSelection = GetRndNo(1, SumOfWeight)
For i = 0 To TotalChoices
    If RandomSelection < ChoiceWeights(i) Then
        Exit For
    Else
        RandomSelection = RandomSelection - ChoiceWeights(i)
    End If
Next i

GetRandomWeightedNo1 = i
End Function


Public Function GetRandomWeightedNo2(ByVal ChoicesAndWeights As Variant) As Variant
'returns random weighted value from array
'ChoicesAndWeights Must be a multidimensional Array. First Column is Choices, Second Column Weights
Dim RandomSelection As Integer
Dim i As Integer
Dim k As Integer
Dim RandomWeightPool() As Variant
ReDim RandomWeightPool(0 To 0)

For i = LBound(ChoicesAndWeights) To UBound(ChoicesAndWeights)
    For k = 0 To ChoicesAndWeights(i, 1)
        RandomWeightPool(UBound(RandomWeightPool)) = ChoicesAndWeights(i, 0)          'Assign the array element
        ReDim Preserve RandomWeightPool(UBound(RandomWeightPool) + 1) 'Allocate next element
    Next k
Next i
ReDim Preserve RandomWeightPool(LBound(RandomWeightPool) To UBound(RandomWeightPool) - 1)  'Deallocate the last, unused eleme
GetRandomWeightedNo2 = RandomWeightPool(GetRndNo(0, UBound(RandomWeightPool)))
End Function



Public Function isBlankOrNull(ByVal TestingValue As Variant)
'returns true if passed value is null or "" reguardless of variable type. I use this so much I almost forget its not a built in VBA Function.
If IsArray(TestingValue) Then
    If IsEmpty(TestingValue) Then
        isBlankOrNull = True
    Else
        isBlankOrNull = False
    End If
Else
    If IsNull(TestingValue) Or TestingValue = "" Then
        isBlankOrNull = True
    Else
        isBlankOrNull = False
    End If
End If
End Function


Public Function IsFileOpen(FileName As String)
'verifies if file is actively open on this or another computer.
    Dim ff As Long, ErrNo As Long
    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: Error ErrNo
    End Select
End Function



Public Function IsNumberKeyPress(ByVal KeyCode As Integer, Optional ByVal Shift As Integer) As Boolean
'verifies if number key on number key row or numberpad is pushed. Usefill in certain situations to only allow number key push on number only fields.
IsNumberKeyPress = False
If (Shift = 2 And KeyCode = 86) Then
    IsNumberKeyPress = True
Else
    Dim KeyAllowList As Variant
    KeyAllowList = Array(vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9)
    Dim i
    For i = LBound(KeyAllowList) To UBound(KeyAllowList)
        If KeyAllowList(i) = KeyCode Then
            IsNumberKeyPress = True
        End If
    Next i
End If
End Function


Public Function IsProcessRunning(process As String) 
'use to check active windows processes
'there seems to be a delay when using this. Seems to be polling a windows database. Sometimes its not initially accurate.
'if you need info that quicker then 5 seconds do not use this.
    Dim objList As Object
    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If
End Function


Public Function openfile(FileNamePath As String) As Boolean
'Open file if exists
If (FileExists(FileNamePath)) Then
    Application.FollowHyperlink FileNamePath
    openfile = True
Else
    MsgBox ("File not found")
    openfile = False
End If

End Function


'required for pause and wait functions remove if not used.
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds as Long) 'MS Office 32 Bit
#End If

Public Function Pause(NumberOfSeconds As Variant) As Boolean
'Waits X time in seconds
'Works even over midnight
'better than using DoEvents which eats up all the CPU cycles
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' after midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop
    Pause = 1
Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Pause = 0
    GoTo Exit_GoTo
End Function


Public Function RandomNumberBetween(ByVal lLowerVal As Double, ByVal lUpperVal As Double, Optional bInclVals As Boolean = True) As Double
'returns random number between uper and lower passed values.
'Max number 999,999,999,999,999

    Dim lTmp As Long
 
    If lLowerVal > lUpperVal Then
        lTmp = lLowerVal
        lLowerVal = lUpperVal
        lUpperVal = lTmp
    End If
 
    If bInclVals = False Then
        lLowerVal = lLowerVal + 1
        lUpperVal = lUpperVal - 1
    End If

    RandomNumberBetween = Int((lUpperVal - lLowerVal + 1) * Rnd + lLowerVal)
End Function



Public Function TransposeArray(myarray As Variant) As Variant
'swaps column and rows of an array
Dim X As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function



'required for pause and wait functions remove if not used.
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds as Long) 'MS Office 32 Bit
#End If

Public Function WaitForTime(datDate As Date)
'Waits until the specified date and time
'better than using DoEvents which eats up all the CPU cycles
  Do
    Sleep 100
    DoEvents
  Loop Until Now >= datDate
End Sub
