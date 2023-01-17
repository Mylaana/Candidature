Attribute VB_Name = "lib_Import"
Option Compare Database
Option Explicit

Public Sub ImportExcelSpreadsheet(fileName As String, tablename As String)
On Error GoTo BadFormat
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Sample", fileName, True

Exit Sub

End Sub

Sub ImportFile(ByVal sImportName As String, Optional ByVal sImportDate As String = "AAAAMMDD", Optional ByVal ImportRule As String = "", Optional ByVal ClearTableOnImport As Boolean = True)


Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xldata As Excel.Range
Dim xlPathFull As String
Dim vTemp As Variant
Dim rsd, rsdRule As DAO.Recordset
Dim sSQL As String
Dim xlBookName, xlPath As String
Dim i, j, k As Long
Dim theSheet As Worksheet

'this sub allows you to either :
' - load 1 excel file with the browser filling N table from N worksheets in the single file (default mode, all import rules without filename/filepath will be considered to be from the same browsed file)
' - load X excel files with filled book/path names, filling N tables from N worksheets of any distribution between X files

If sImportDate = "AAAAMMDD" Then sImportDate = ""

sSQL = "SELECT SYS_Import.* FROM SYS_Import where sys_import.importname = " & Entrecote(sImportName) & " and bimport = true ORDER BY SYS_Import.ImportName"
Set rsd = CurrentDb.OpenRecordset(sSQL)

If rsd.EOF Then GoTo SubCleaning
rsd.MoveLast: rsd.MoveFirst

If rsd("FileType") = "Excel" Then
    Set xlapp = New Excel.Application
End If

For k = 1 To rsd.RecordCount
    If rsd("sourcefilename") & "" <> "" And rsd("sourcefilepath") & "" <> "" Then
    'case 1 : the importTable is filled with file coordinate
        
        If rsd("sourcefilename") & "" = xlBookName & "" And Replace(rsd("sourcefilepath") & "\", "\\", "\") = xlPath & "" Then
            'same file, we do nothing
        Else
            'different file, we close the opened one if exists and open the new one
            If xlBookName & "" <> "" Then xlBook.Close False
            
            xlBookName = rsd("sourcefilename")
            xlPath = Replace(rsd("sourcefilepath") & "\", "\\", "\")
            Set xlBook = xlapp.Workbooks.Open(xlPath & xlBookName, , xlReadOnly)
        End If
    
    
    Else
    'case 2 : no coordinates so we use MsoFile to browse for the file
        
        If xlBookName & "" = "" Then
            xlPathFull = GetFilePathBrowser 'xlapp.GetOpenFilename("Excel files (*.xls; *.xlsx; *.xlsm; ),*.xls; *.xlsx; *.xlsm; ", , , , False)
            If CStr(xlPathFull) = CStr(Empty) Then
                MsgBox "No file selected, please select one"
                GoTo SubCleaning
            End If
            
            xlapp.Workbooks.Open (xlPathFull)
            Set xlBook = xlapp.ActiveWorkbook
            xlBookName = xlapp.ActiveWorkbook.Name
            xlPath = xlBook.Path 'Split(xlBook.Path, "\")(UBound(Split(xlBook.Path, "\")))
            xlapp.Visible = True
        End If
    End If
    
    'purge de la table
    If ClearTableOnImport Then
        CurrentDb.Execute "DELETE * FROM " & rsd("desttablename")
    End If
    
    'copie des données
    For Each theSheet In xlBook.Worksheets
        If Replace(Nz(rsd("sourcesheetname"), ""), "*", "") <> rsd("sourcesheetname") Then
            If Left(rsd("sourcesheetname"), 1) = "*" And Right(rsd("sourcesheetname"), 1) = "*" Then
                'searching for a sheet name that contains the sourcesheetname
                If UCase(Replace(theSheet.Name, Replace(rsd("sourcesheetname"), "*", ""), "")) <> UCase(theSheet.Name) Then
                    vTemp = xlapp.Workbooks(xlBook.Name).Worksheets(theSheet.Name).UsedRange
                    Exit For
                End If
            ElseIf Left(rsd("sourcesheetname"), 1) <> "*" And Right(rsd("sourcesheetname"), 1) = "*" Then
                'searching for a sheet name that begins with the sourcesheetname
                If UCase(Left(theSheet.Name, Len(Replace(rsd("sourcesheetname"), "*", "")))) = UCase(Replace(rsd("sourcesheetname"), "*", "")) Then
                    vTemp = xlapp.Workbooks(xlBook.Name).Worksheets(theSheet.Name).UsedRange
                    Exit For
                End If
            ElseIf Left(rsd("sourcesheetname"), 1) = "*" And Right(rsd("sourcesheetname"), 1) <> "*" Then
                'searching for a sheet name that ends with the sourcesheetname
                If UCase(Right(theSheet.Name, Len(Replace(rsd("sourcesheetname"), "*", "")))) = UCase(Replace(rsd("sourcesheetname"), "*", "")) Then
                    vTemp = xlapp.Workbooks(xlBook.Name).Worksheets(theSheet.Name).UsedRange
                    Exit For
                End If
            End If
        Else
                'searching for a sheet name that matches exactly sourcesheetname
                If UCase(Nz(rsd("sourcesheetname"), "")) = UCase(theSheet.Name) Then
                    Set vTemp = Nothing
                    vTemp = xlapp.Workbooks(xlBook.Name).Worksheets(theSheet.Name).UsedRange.Value
                    Exit For
                
                'taking the first worksheet if no name is filled in the import table
                ElseIf Nz(rsd("sourcesheetname"), "") = "" Then
                    Set vTemp = Nothing
                    vTemp = xlapp.Workbooks(xlBook.Name).Worksheets(theSheet.Name).UsedRange.Value
                    Exit For
                End If

        End If
    Next
    
    If IsArray(vTemp) Then
        'Customize fields name
        If ImportRule & "" <> "" Then
            Set rsdRule = CurrentDb.OpenRecordset("select * from SYS_ImportRule where importrule = " & Entrecote(ImportRule))
            If Not rsdRule.EOF Then rsdRule.MoveLast: rsdRule.MoveFirst
                    
            For i = LBound(vTemp, 2) To UBound(vTemp, 2)
                rsdRule.MoveFirst
                For j = 1 To rsdRule.RecordCount
                    If vTemp(1, i) = rsdRule("SourceFieldName") Then
                        vTemp(1, i) = rsdRule("tableFieldName")
                    End If
                    rsdRule.MoveNext
                Next j
            Next i
        End If
            
        'remplissage de la table
        Call FillTable(vTemp, rsd("desttablename"))
        vTemp = Null
    End If
    
    rsd.Edit
    rsd("importdate") = sImportDate
    rsd("importuser") = Environ$("Username")
    'rsd("sourcefilepath") = xlPath
    'rsd("sourcefilename") = xlBookName
    rsd.Update
    
    rsd.MoveNext
Next k

xlBook.Close False
xlapp.Quit
Set xlapp = Nothing

'update Param table
CurrentDb.Execute "UPDATE SYS_PARAM SET SYS_PARAM.TheValue = " & Entrecote(Right(sImportDate, 2) & "/" & Left(Right(sImportDate, 4), 2) & "/" & Left(sImportDate, 4)) & " WHERE SYS_PARAM.TheParam='Actuals DD/MM/YYYY' "


DoEvents
MsgBox "End of treatment"


SubCleaning:
Set xlapp = Nothing



Exit Sub


ErrorHandler:

MsgBox "Could not import data"
GoTo SubCleaning

End Sub

Private Sub FillTable(ByVal vTemp As Variant, ByVal sTableName As String)

Dim rsd As DAO.Recordset
Dim i, j, k As Long


Set rsd = Application.CurrentDb.OpenRecordset("select * from " & sTableName)

If Not rsd.EOF Then rsd.MoveLast: rsd.MoveFirst

For i = LBound(vTemp) To UBound(vTemp) - 1
    rsd.AddNew
    'boucle sur les colonnes de vTemp
        
    For j = LBound(vTemp, 2) To UBound(vTemp, 2)
        'recherche du nom de colonne correspondant dans la table
        For k = 0 To rsd.Fields.Count - 1
            If rsd.Fields(k).Name = vTemp(1, j) Then
                'copie de la valeur
                If IsError(vTemp(i + 1, j)) Then
                    rsd(k) = Empty
                Else
                    On Error Resume Next
                    rsd(k) = Nz(vTemp(i + 1, j), "")
                    On Error GoTo 0
                    If IsNull(rsd(k)) Then rsd(k) = Empty
                End If
                Exit For
            End If
        Next
        
    Next
    rsd.Update
Next


End Sub
