Attribute VB_Name = "lib_Actuals"
Option Compare Database
Option Explicit

Sub Fill_T10_Actuals()

Dim i, j, k As Long
Dim rsdSrc, rsdDest As DAO.Recordset
Dim vTableList As Variant
Dim rsdSTR As String

ReDim vTableList(1 To 3)
vTableList(1) = "T11_Actuals_US"
vTableList(2) = "T12_Actuals_PAR"
vTableList(3) = "T13_Actuals_SIN"

'clear T10 & parametrage
CurrentDb.Execute ("delete * from T10_Actuals")
Set rsdDest = Application.CurrentDb.OpenRecordset("select * from T10_Actuals")
If Not (rsdDest.EOF) Then rsdDest.MoveLast: rsdDest.MoveFirst

'copie des tables secteurs dans la table main
For k = LBound(vTableList) To UBound(vTableList)
    Set rsdSrc = Application.CurrentDb.OpenRecordset("select * from " & vTableList(k))
    
    If Not rsdSrc.EOF Then rsdSrc.MoveLast: rsdSrc.MoveFirst
    
    Do Until rsdSrc.EOF
        If rsdSrc("line_total_excl_tax") & "" <> "" Then
            rsdDest.AddNew
            For j = 0 To rsdDest.Fields.Count - 1
                Select Case rsdDest.Fields(j).Name
                    Case "unit_price_excl_tax", "line_total_incl_tax", "tax_amount", "line_total_excl_tax", "quantity"
                        'numeric value
                        If IsNumeric(Nz(rsdSrc(rsdDest.Fields(j).Name), 0)) Then
                            rsdDest.Fields(j) = Nz(rsdSrc(rsdDest.Fields(j).Name), 0)
                        Else
                            rsdDest.Fields(j) = Replace(Replace(Nz(rsdSrc(rsdDest.Fields(j).Name), 0), ",", ""), ".", ",")
                        End If
                    
                    Case "tax_rate"
                        'percentage
                        rsdDest.Fields(j) = Replace(Replace(Replace(Nz(rsdSrc(rsdDest.Fields(j).Name), 0), "%", ""), ",", ""), ".", ",") / 100
                    
                    Case "Estimated_billing_date", "service_end_date", "service_start_date"
                        'dates
                        rsdDest.Fields(j) = CDate(Replace(Nz(rsdSrc(rsdDest.Fields(j).Name), 0), "??", 0))
                                        
                    Case "service_code"
                        rsdDest.Fields(j) = rsdSrc(rsdDest.Fields(j).Name)
                            If rsdDest.Fields(j) & "" = "" Then
                                rsdDest.Fields(j) = "ERROR"
                            End If
                        
                    Case "isadjustment"
                        'do nothing
                    Case Else
                        'autre
                        rsdDest.Fields(j) = (rsdSrc(rsdDest.Fields(j).Name))
                End Select
                
            Next
            rsdDest.Update
        End If
        rsdSrc.MoveNext
    Loop
    

    Set rsdSrc = Nothing
Next

'CurrentDb.Execute "UPDATE_T10_CSM"


End Sub

Sub Fill_T21_Synthesis_Detail(sTblSource As String, sTblDest As String)


Dim i, j, k, a As Long
Dim rsdSrc, rsdDest As DAO.Recordset
Dim vTableList, vMonthlyAmount As Variant
Dim sSQL As String
Dim iNbMois As Long
Dim sServStartDateCorr, sServEndDateCorr, invoiceDateCorr As String
Dim bFieldCheck As Boolean
Dim dTotalSpreadPercentage, dTotalSpread As Double

Dim dTotalSpreadSansDate As Double

'clear T21 & settings
CurrentDb.Execute ("delete * from " & sTblDest)
Set rsdDest = Application.CurrentDb.OpenRecordset("select * from " & sTblDest)
If Not (rsdDest.EOF) Then rsdDest.MoveLast: rsdDest.MoveFirst

'T10 Loading
sSQL = "SELECT " & sTblSource & ".*, REF_Service_Tradename.Business_Type, REF_Service_Tradename.Business_SubType, REF_Service_Tradename.Product_Type " & _
        "FROM REF_Service_Tradename RIGHT JOIN " & sTblSource & " ON (REF_Service_Tradename.service_tradename = " & sTblSource & ".service_tradename) AND (REF_Service_Tradename.type = " & sTblSource & ".type) AND (REF_Service_Tradename.service_code = " & sTblSource & ".service_code)"

'Debug.Print sSQL
Set rsdSrc = Application.CurrentDb.OpenRecordset(sSQL)
If Not (rsdSrc.EOF) Then rsdSrc.MoveLast: rsdSrc.MoveFirst


Do Until rsdSrc.EOF
    If Not (rsdSrc("document_date") & "" = "" And Nz(rsdSrc("line_total_excl_tax"), 0) = 0 And Nz(rsdSrc("document_date"), 0) = 0) Then
        'start/end/invoice date calculation before detailed lines creation
        If Nz(rsdSrc("estimated_billing_date"), "") & "" <> "" And CDate(Nz(rsdSrc("estimated_billing_date"), 0)) <> CDate(0) Then
            invoiceDateCorr = rsdSrc("estimated_billing_date")
        Else
            invoiceDateCorr = CDate(Nz(rsdSrc("document_date"), 0))
        End If
        
        If rsdSrc("service_start_date") & "" <> "" And CDate(Nz(rsdSrc("service_start_date"), 0)) <> CDate(0) Then
            sServStartDateCorr = rsdSrc("service_start_date")
        Else
            sServStartDateCorr = invoiceDateCorr
        End If
        
        If Nz(rsdSrc("service_end_date"), "") & "" <> "" And CDate(Nz(rsdSrc("service_end_date"), 0)) <> CDate(0) Then
            sServEndDateCorr = rsdSrc("service_end_date")
        Else
            If rsdSrc("business_type") = "recurrent" Then
                dTotalSpreadSansDate = dTotalSpreadSansDate + rsdSrc("line_total_excl_tax")
                sServEndDateCorr = DateAdd("m", 12, sServStartDateCorr)
            Else
                sServEndDateCorr = sServStartDateCorr
            End If
        End If
        
        'Autocorrection des dates fausses
        If CDate(sServEndDateCorr) < CDate(sServStartDateCorr) Then
            If rsdSrc("business_type") = "recurrent" Then
                sServEndDateCorr = DateAdd("y", 12, sServStartDateCorr)
            Else
                sServEndDateCorr = sServStartDateCorr
            End If
            dTotalSpreadSansDate = dTotalSpreadSansDate + rsdSrc("line_total_excl_tax")
    
        End If
        
        iNbMois = (Year(sServEndDateCorr) - Year(sServStartDateCorr)) * 12 + (Month(sServEndDateCorr) - Month(sServStartDateCorr)) + 1
        
        'Calculation of most dates related fields and monthly spread the invoice amount
        '1 = start month | 2 = end month | 3 = Start service corrected on month |  4 = end service corrected on month | 5 = percentage to spread | 6 = amount
        ReDim vMonthlyAmount(1 To iNbMois, 1 To 6)
        dTotalSpreadPercentage = 0
        For i = LBound(vMonthlyAmount) To UBound(vMonthlyAmount)
            'start month date
            vMonthlyAmount(i, 1) = DateSerial(Year(sServStartDateCorr), Month(sServStartDateCorr) + i - 1, 1)
            'end month date
            vMonthlyAmount(i, 2) = DateSerial(Year(sServStartDateCorr), Month(sServStartDateCorr) + i, 1) - 1
            
            'start serv corr on month
            If Year(sServStartDateCorr) = Year(vMonthlyAmount(i, 1)) And Month(sServStartDateCorr) = Month(vMonthlyAmount(i, 1)) Then
                vMonthlyAmount(i, 3) = sServStartDateCorr
            Else
                vMonthlyAmount(i, 3) = vMonthlyAmount(i, 1)
            End If
            
            'end service date on month
            If Year(sServEndDateCorr) = Year(vMonthlyAmount(i, 2)) And Month(sServEndDateCorr) = Month(vMonthlyAmount(i, 2)) Then
                vMonthlyAmount(i, 4) = sServEndDateCorr
            Else
                vMonthlyAmount(i, 4) = vMonthlyAmount(i, 2)
            End If
            
            'percentage of the month covering the service
            If i < iNbMois Then
                vMonthlyAmount(i, 5) = (CDate(vMonthlyAmount(i, 4)) - CDate(vMonthlyAmount(i, 3)) + 1) / (CDate(vMonthlyAmount(i, 2)) - CDate(vMonthlyAmount(i, 1)) + 1)
                dTotalSpreadPercentage = dTotalSpreadPercentage + vMonthlyAmount(i, 5)
            Else
                vMonthlyAmount(i, 5) = 1 - dTotalSpreadPercentage
                dTotalSpreadPercentage = dTotalSpreadPercentage + 1
            End If
        Next
        
        'spread amount
        dTotalSpread = 0
        For i = LBound(vMonthlyAmount) To UBound(vMonthlyAmount)
            'spread amount
            If i < iNbMois Then
                vMonthlyAmount(i, 6) = rsdSrc.Fields("line_total_excl_tax") * vMonthlyAmount(i, 5) / dTotalSpreadPercentage
                dTotalSpread = dTotalSpread + vMonthlyAmount(i, 6)
            Else
                vMonthlyAmount(i, 6) = rsdSrc.Fields("line_total_excl_tax") - dTotalSpread
                'If dTotalSpread + rsdSrc.Fields("line_total_excl_tax") * vMonthlyAmount(i, 5) / dTotalSpreadPercentage <> rsdSrc.Fields("line_total_excl_tax") Then Stop
            End If
        Next
        
        
        'looping on number of invoced months
        For i = 1 To iNbMois
            rsdDest.AddNew
            For j = 0 To rsdDest.Fields.Count - 1
                Select Case rsdDest.Fields(j).Name
                    Case "invoice date corrected"
                        rsdDest.Fields(j) = CDate(invoiceDateCorr)
                        
                    Case "service_start_date corrected"
                            rsdDest.Fields(j) = sServStartDateCorr
                        
                    Case "service_end_date corrected"
                            rsdDest.Fields(j) = sServEndDateCorr
    
                    
                    Case "Start month date"
                        rsdDest.Fields(j) = vMonthlyAmount(i, 1)
    
                    Case "end month date"
                        rsdDest.Fields(j) = vMonthlyAmount(i, 2)
                        
                    Case "service begining date on month"
                        rsdDest.Fields(j) = vMonthlyAmount(i, 3)
                        
                        
                    Case "service ending date on month"
                        rsdDest.Fields(j) = vMonthlyAmount(i, 4)
    
                    Case "nb days of svc total"
                        rsdDest.Fields(j) = CDate(sServEndDateCorr) - CDate(sServStartDateCorr) + 1
                        
                    Case "nb days of svc on month"
                        rsdDest.Fields(j) = rsdDest.Fields("service ending date on month") - rsdDest.Fields("service begining date on month") + 1
    
                    Case "revenue excl taxes per day in transaction currency"
                        rsdDest.Fields(j) = rsdDest.Fields("line_total_excl_tax") / rsdDest.Fields("nb days of svc total")
    
                    Case "invoice Year"
                        rsdDest.Fields(j) = Year(invoiceDateCorr)
                        
                    Case "invoice month"
                        rsdDest.Fields(j) = Month(invoiceDateCorr)
                        
                    Case "service Year"
                        rsdDest.Fields(j) = Year(vMonthlyAmount(i, 3))
                        
                    Case "service month"
                        rsdDest.Fields(j) = Month(vMonthlyAmount(i, 3))
                        
                    Case "TheCurrency"
                        rsdDest.Fields(j) = rsdSrc("theCurrency")
                    
'                    Case "service_code"
'                        rsdDest.Fields(j) = rsdSrc("service_code")
'                        If rsdDest.Fields(j) & "" = "" Then
'                            rsdDest.Fields(j) = "ERROR"
'                        End If
    '                Case "days of service to cut date"
    
    '                Case "Days before cut off"
    
    '                Case "days after cut off"
    
                    Case "total local curr on month"
                        rsdDest.Fields(j) = vMonthlyAmount(i, 6)
    
    '                Case "total Euro on month"
                    
                    Case Else
                        bFieldCheck = False
                        On Error Resume Next
                            bFieldCheck = rsdDest.Fields(j).Name = rsdSrc(rsdDest.Fields(j).Name).Name
                        On Error GoTo 0
                            If bFieldCheck Then
                                If IsNumeric(Nz(rsdSrc(rsdDest.Fields(j).Name), "")) Then
                                    'other numeric
                                    rsdDest.Fields(j) = rsdSrc(rsdDest.Fields(j).Name)
                                Else
                                    'other non numeric
                                    If Left(rsdSrc(rsdDest.Fields(j).Name), 18) = "FORECAST_SUPPORT_L" Then
                                    'Stop
                                    End If
                                    rsdDest.Fields(j) = (rsdSrc(rsdDest.Fields(j).Name))
                                End If
                            End If
                End Select
    
            Next j
            rsdDest.Update
        Next
    End If
    rsdSrc.MoveNext
    DoEvents
Loop




End Sub

Sub TreatVRactuals()


CurrentDb.Execute "DELETE_T19"
CurrentDb.Execute "ADD_VR_CustommerTransaction"
CurrentDb.Execute "ADD_VR_RRFollowup_RRHisto"
CurrentDb.Execute "ADD_VR_RRFollowup_NRRHisto"
CurrentDb.Execute "ADD_VR_RRFollowup_RR"
CurrentDb.Execute "ADD_VR_RRFollowup_NRR"
CurrentDb.Execute "ADD_VR_Provisional"
CurrentDb.Execute "ADD_T19_Adjustments_VR"
CurrentDb.Execute "UPDATE_T19_INTERCO"
CurrentDb.Execute "UPDATE_T19_PriceExclTax"
CurrentDb.Execute "UPDATE_VR_ClientNameFromCustTransac"
CurrentDb.Execute "UPDATE_T19_IDProject_RRFollowup"

End Sub


