Attribute VB_Name = "lib_Gene"
Option Compare Database
Option Explicit

Function Entrecote(ByVal sMyString)

'sMyString = Replace(sMyString, "''", "'")
If Replace(sMyString, "'", "") = sMyString Then
    Entrecote = "'" & sMyString & "'"
Else
    Entrecote = Replace(Nz(sMyString, ""), "'", "''")
End If

End Function
Function Fieldexists(tablename As String, fieldname As String) As Boolean

Dim exists As Boolean

   exists = False
   On Error Resume Next
   exists = CurrentDb.TableDefs(tablename).Fields(fieldname).Name = fieldname

   Fieldexists = exists
   
End Function
Function GetSysParam(ParamName As String)

GetSysParam = DLookup("Thevalue", "SYS_Param", "TheParam=" & Entrecote(ParamName))

End Function
Function GetDateYYYYMMDDToSerial(theDate As String)

GetDateYYYYMMDDToSerial = DateSerial(Left(theDate, 4), Right(Left(theDate, 6), 2), Right(theDate, 2))

End Function
Function GetQueryResult(TheQueryName As String)


Dim rsd As DAO.Recordset
Dim strSQL As String

strSQL = "SELECT * FROM " & TheQueryName
Set rsd = CurrentDb.OpenRecordset(strSQL)
' new code:
GetQueryResult = rsd.Fields(0).Value
rsd.Close
Set rsd = Nothing


End Function

Function GetFilePathBrowser(Optional ByVal FileTypeName As String = "Excel", Optional ByVal FileType As String = "*.xls; *.csv; *.xlsx; *.xlsm", Optional ByVal MultiselectAllowed As Boolean = False)

Dim fd As FileDialog

'Create a FileDialog object as a File Picker dialog.
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'Declare a variable to contain the path
'of each selected item. Even though the path is aString,
'the variable must be a Variant because For Each...Next
'routines only work with Variants and Objects.
Dim vrtSelectedItem As Variant

'Use a With...End With block to reference the FileDialog object.
With fd

    'Add a filter that includes XL files and make it the first item in the list.
    .Filters.Add FileTypeName, FileType, 1
    
    'forbid multiselect
    .AllowMultiSelect = MultiselectAllowed
    
    'Use the Show method to display the File Picker dialog box and return the user's action.
    'If the user presses the button...
    If .Show = -1 Then
        
        'Step through each string in the FileDialogSelectedItems collection.
        For Each vrtSelectedItem In .SelectedItems
        
            'vrtSelectedItem is aString that contains the path of each selected item.
            'You can use any file I/O functions that you want to work with this path.
            'This example displays the path in a message box.
            GetFilePathBrowser = vrtSelectedItem
        
        Next vrtSelectedItem
        
        'If the user presses Cancel...
        Else
        GetFilePathBrowser = Empty
        'MsgBox "No file selected"
        
    End If
End With

'Set the object variable to Nothing.
Set fd = Nothing

End Function
Function GetSplitArray(ByVal vArrayToSplit As Variant, Optional ByVal ArraySize As Long = 50000)

Dim vTemp As Variant
Dim vTempGroup As Variant
Dim i, j, k, vLine, vSize As Long

vSize = UBound(vArrayToSplit) \ ArraySize
If UBound(vArrayToSplit) Mod ArraySize > 0 Then vSize = vSize + 1

ReDim vTempGroup(1 To vSize)

vLine = 1
For k = 1 To vSize
    If k = vSize Then
        ReDim vTemp(LBound(vArrayToSplit) To UBound(vArrayToSplit) Mod ArraySize, LBound(vArrayToSplit, 2) To UBound(vArrayToSplit, 2))
    Else
        ReDim vTemp(LBound(vArrayToSplit) To ArraySize, LBound(vArrayToSplit, 2) To UBound(vArrayToSplit, 2))
    End If
    
    For i = LBound(vTemp) To UBound(vTemp)
        For j = LBound(vTemp, 2) To UBound(vTemp, 2)
            vTemp(i, j) = vArrayToSplit(vLine, j)
        Next j
        vLine = vLine + 1
    Next i
    
    vTempGroup(k) = vTemp
Next

GetSplitArray = vTempGroup

End Function

Function GetQuerySQL(TheQueryName As String) As String

Dim QD As DAO.QueryDef
 
Set QD = CurrentDb.QueryDefs(TheQueryName)
GetQuerySQL = QD.SQL
 
End Function

