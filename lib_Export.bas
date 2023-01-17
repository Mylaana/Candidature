Attribute VB_Name = "lib_Export"
Option Compare Database
Option Explicit

Sub exportExcelTemplate(sExportName As String)


Dim i, j, k As Long
Dim rsdQuery, rsdExport As DAO.Recordset
Dim sSQL As String

Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xldata As Excel.Range
Dim xlPath As String
Dim vTemp, vTempSub, vFilters As Variant


sSQL = "SELECT SYS_Export.* from SYS_Export where SYS_Export.isactive = true and SYS_Export.Exportname = " & Entrecote(sExportName) & "order by sys_export.ID_Export"
Set rsdExport = CurrentDb.OpenRecordset(sSQL)

If Not rsdExport.EOF Then
    rsdExport.MoveLast: rsdExport.MoveFirst
Else
    MsgBox "cannot export the requested data"
    Exit Sub
End If

xlPath = Application.CurrentProject.Path & "\TEMPLATE\"
Set xlapp = New Excel.Application
xlapp.Workbooks.Open (xlPath & rsdExport("templatename"))
Set xlBook = xlapp.ActiveWorkbook


For k = 1 To rsdExport.RecordCount
    Set rsdQuery = Nothing
    
    If rsdExport("isexportresult") = False Then
        'exécution query/sub
        
        If rsdExport("thequery") & "" <> "" Then CurrentDb.Execute rsdExport("thequery")
        If rsdExport("thesub") & "" <> "" Then
            Application.Run rsdExport("thesub"), xlapp
        End If
    Else
        sSQL = GetQuerySQL(rsdExport("thequery"))
        
        'apply filters
        If rsdExport("Filters") & "" <> "" Then
            vFilters = Split(rsdExport("Filters"), ",")
            
            For i = LBound(vFilters) To UBound(vFilters)
                sSQL = (Replace(sSQL, Trim(Split(vFilters(i), "=")(0)), Trim(Split(vFilters(i), "=")(1))))
            Next i
        End If
        
        Set rsdQuery = CurrentDb.OpenRecordset(sSQL)
        Set vTemp = Nothing
        
        
        If Not rsdQuery.EOF Then
            rsdQuery.MoveLast: rsdQuery.MoveFirst
            ReDim vTemp(1 To rsdQuery.RecordCount + 1, 1 To rsdQuery.Fields.Count)
            For i = LBound(vTemp) To UBound(vTemp)
                For j = LBound(vTemp, 2) To UBound(vTemp, 2)
                    If i = LBound(vTemp) Then
                        vTemp(i, j) = rsdQuery.Fields(j - 1).Name
                    Else
                        vTemp(i, j) = rsdQuery(j - 1)
    
                    End If
                Next
                If i <> LBound(vTemp) Then rsdQuery.MoveNext
            Next
        End If
        
        'past data
        With xlapp.Workbooks(Split(xlBook.Name, ".")(0)).Worksheets(CStr(rsdExport("destsheetname")))
            If .UsedRange.Rows.Count > 1 Then
                'removing filter before deleting data
                On Error Resume Next
                    .AutoFilter.ShowAllData
                On Error GoTo 0
                .Rows("2:" & .UsedRange.Rows.Count).Delete
            End If
            'Debug.Print CStr(rsdExport("destsheetname"))

            If Not rsdQuery.RecordCount = 0 Then
                vTemp = GetSplitArray(vTemp)
                
                For i = LBound(vTemp) To UBound(vTemp)
                DoEvents
                    vTempSub = vTemp(i)
                    .Range(.Cells(.UsedRange.Rows.Count, 1), .Cells(.UsedRange.Rows.Count + UBound(vTempSub) - 1, UBound(vTempSub, 2))) = vTempSub
                Next i
            End If
        End With

    End If
    rsdExport.MoveNext
Next k

'Unload Forms(Form_ExportMngr.Name)
xlapp.WindowState = xlMaximized
xlapp.Visible = True

MsgBox "End of Treatment"

End Sub

Sub exportExcelWithoutTemplate(sExportName As String)

Dim i, j, k As Long
Dim rsdQuery, rsdExport As DAO.Recordset
Dim sSQL As String

Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xldata As Excel.Range
Dim xlPath As String
Dim vTemp, vTempSub As Variant
Dim xlSheetName As String
Dim theRange As Variant


sSQL = "SELECT SYS_Export.* from SYS_Export where SYS_Export.isactive = true and SYS_Export.Exportname = " & Entrecote(sExportName) & "order by sys_export.ID_Export"
Set rsdExport = CurrentDb.OpenRecordset(sSQL)

If Not rsdExport.EOF Then
    rsdExport.MoveLast: rsdExport.MoveFirst
Else
    MsgBox "cannot export the requested data"
    Exit Sub
End If

Set xlapp = New Excel.Application
xlapp.Workbooks.Add
Set xlBook = xlapp.ActiveWorkbook

For k = 1 To rsdExport.RecordCount
    Set rsdQuery = Nothing
    If rsdExport("isexportresult") = False Then
        'exécution query/sub
        If rsdExport("thequery") & "" <> "" Then CurrentDb.Execute rsdExport("thequery")
        If rsdExport("thesub") & "" <> "" Then
            Application.Run rsdExport("thesub"), xlapp
        End If
    End If
    Set rsdQuery = CurrentDb.OpenRecordset(rsdExport("thequery"))

    If Not rsdQuery.EOF Then rsdQuery.MoveLast: rsdQuery.MoveFirst
        
    ReDim vTemp(1 To rsdQuery.RecordCount + 1, 1 To rsdQuery.Fields.Count)
    For i = LBound(vTemp) To UBound(vTemp)
        For j = LBound(vTemp, 2) To UBound(vTemp, 2)
            If i = 1 Then
                vTemp(i, j) = rsdQuery.Fields(j - 1).Name
            Else
                vTemp(i, j) = rsdQuery(j - 1)
            End If
        Next
        If i <> 1 Then rsdQuery.MoveNext
    Next
        
    'sheet name
    xlSheetName = rsdExport("DestsheetName") & ""
    If xlSheetName & "" = "" Then xlSheetName = rsdExport("thequery")
    If k = 1 Then
        Set xlSheet = xlapp.Workbooks(xlBook.Name).ActiveSheet
    Else
        xlapp.Workbooks(xlBook.Name).Worksheets.Add
        Set xlSheet = xlapp.Workbooks(xlBook.Name).ActiveSheet
    End If
    xlSheet.Name = xlSheetName
    
    vTemp = GetSplitArray(vTemp)
    
    For i = LBound(vTemp) To UBound(vTemp)
    DoEvents
        vTempSub = vTemp(i)
        With xlapp.Workbooks(Split(xlBook.Name, ".")(0)).Worksheets(xlSheet.Name)
            If i = 1 Then
                .Range(.Cells(.UsedRange.Rows.Count, 1), .Cells(.UsedRange.Rows.Count + UBound(vTempSub) - 1, UBound(vTempSub, 2))) = vTempSub
            Else
                .Range(.Cells(.UsedRange.Rows.Count + 1, 1), .Cells(.UsedRange.Rows.Count + UBound(vTempSub), UBound(vTempSub, 2))) = vTempSub
            End If
            
            'Autosize columns
            .Range(.Cells(.UsedRange.Rows.Count, 1), .Cells(.UsedRange.Rows.Count + UBound(vTempSub) - 1, UBound(vTempSub, 2))).EntireColumn.AutoFit
            
        End With
    Next i
        
    rsdExport.MoveNext
Next k

xlapp.Visible = True
MsgBox "End of Treatment"

End Sub
Sub SetRenewalForecastCutOffDate(Optional xlapp As Excel.Application)

Dim sCutoff As String
Dim sDefaultDate As String
Dim rsdQuery As DAO.Recordset

Set rsdQuery = CurrentDb.OpenRecordset("SELECT SYS_PARAM.TheValue FROM SYS_PARAM WHERE SYS_PARAM.TheParam ='Actuals DD/MM/YYYY'")
sDefaultDate = rsdQuery("thevalue")

sCutoff = InputBox("Enter a cut off date please. Format must be 'DD/MM/AAAA', by default the query will use " & sDefaultDate & " if no date is entered", "Set cut off date")
If sCutoff & "" = "" Then sCutoff = sDefaultDate
If IsDate(sCutoff) = False Then
    MsgBox ("Invalid date entered, the system will use " & sDefaultDate)
    sCutoff = sDefaultDate
End If

Set rsdQuery = CurrentDb.OpenRecordset("SELECT SYS_PARAM.TheValue FROM SYS_PARAM WHERE SYS_PARAM.TheParam ='ForecastRenewalCutOff'")
rsdQuery.Edit
rsdQuery("thevalue") = sCutoff
rsdQuery.Update

End Sub
Sub SetFAECutOffDate(Optional xlapp As Excel.Application)

Dim sCutoff As String
Dim sDefaultDate As String
Dim rsdQuery As DAO.Recordset

Set rsdQuery = CurrentDb.OpenRecordset("SELECT SYS_PARAM.TheValue FROM SYS_PARAM WHERE SYS_PARAM.TheParam ='CutOffFAE'")
sDefaultDate = rsdQuery("thevalue")

sCutoff = InputBox("Enter a cut off date please. Format must be 'DD/MM/AAAA', by default the query will use " & sDefaultDate & " if no date is entered", "Set cut off date")
If sCutoff & "" = "" Then sCutoff = sDefaultDate
If IsDate(sCutoff) = False Then
    MsgBox ("Invalid date entered, the system will use " & sDefaultDate)
    sCutoff = sDefaultDate
End If

Set rsdQuery = CurrentDb.OpenRecordset("SELECT SYS_PARAM.TheValue FROM SYS_PARAM WHERE SYS_PARAM.TheParam ='CutOffFAE'")
rsdQuery.Edit
rsdQuery("thevalue") = sCutoff
rsdQuery.Update

End Sub

Sub ExportRevenuePerProject(Optional xlapp As Excel.Application)

Dim vTemp As Variant
Dim xlSheet As Worksheet
Dim i, iColumn As Long

iColumn = 4

For Each xlSheet In xlapp.ActiveWorkbook.Worksheets
    If Replace(xlSheet.Name & "", "DATA Project", "") <> xlSheet.Name & "" Then
        Set vTemp = Nothing
        
        With xlapp.ActiveWorkbook.Worksheets(xlSheet.Name)
            vTemp = .Range(.Cells(2, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count))
    
            For i = LBound(vTemp) To UBound(vTemp)
                vTemp(i, iColumn) = i
            Next
            
            .Range(.Cells(2, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)) = vTemp
        End With
    End If
Next


End Sub


