Attribute VB_Name = "lib_System"
Option Compare Database


Sub UpdateLstQuery()

Dim qdf As QueryDef
Dim rsd As Recordset
Dim sSQL As String

CurrentDb.Execute "DELETE * from LST_Querys"
Set rsd = CurrentDb.OpenRecordset("SELECT * from LST_querys")

For Each qdf In CurrentDb.QueryDefs
    sSQL = ""
    On Error Resume Next
    sSQL = GetQuerySQL(qdf.Name)
    On Error GoTo 0
    
    If Left(sSQL, 6) = "SELECT" Then
        rsd.AddNew
        rsd("queryname") = qdf.Name
        rsd("theSQL") = sSQL
        
        'test t10 & t21
        If Replace(sSQL, "T10", "") <> sSQL Then rsd("HasT10") = -1
        If Replace(sSQL, "T21", "") <> sSQL Then rsd("HasT21") = -1
        
        rsd.Update
    End If
Next qdf

End Sub

