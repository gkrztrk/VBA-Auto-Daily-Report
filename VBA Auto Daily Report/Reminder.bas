Attribute VB_Name = "Reminder"
Dim WBS As Variant
Dim WBS2, WBS3, WBS4, WBS5, WBS6 As String
Dim sourceRow, ReportRow, fnshRow As Integer
Dim dayRange As Integer
Dim status As Double




Sub ReminderReport()
dayRange = 15
sourceRow = 4
ReportRow = Report.Cells(Rows.Count, 1).End(xlUp).Row
fnshRow = Finished.Cells(Rows.Count, 1).End(xlUp).Row

Report.Range("A2:K" & ReportRow) = ""
Finished.Range("A2:I" & fnshRow) = ""
Report.Select

For sourceRow = 4 To 799

    WBS = Tabelle24.Range("G" & sourceRow).Value
    
    Select Case WBS
    
    Case 2:
    
        WBS2 = Tabelle24.Range("H" & sourceRow).Value
    Case 3:
    
        WBS3 = Tabelle24.Range("H" & sourceRow).Value
        
    Case 4:
    
        WBS4 = Tabelle24.Range("H" & sourceRow).Value
    
    Case 5:
    
        WBS5 = Tabelle24.Range("H" & sourceRow).Value
        
    Case 6:
    
        WBS6 = Tabelle24.Range("H" & sourceRow).Value
        
    End Select
    


        If Tabelle24.Range("N" & sourceRow).Value < dayRange And WBS = "A" And _
        Tabelle24.Range("N" & sourceRow).Value >= 0 Then
    
            ReportRow = Report.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            Report.Cells(ReportRow, 1) = WBS2
            Report.Cells(ReportRow, 2) = WBS3
            Report.Cells(ReportRow, 3) = WBS4
            Report.Cells(ReportRow, 4) = WBS5
            Report.Cells(ReportRow, 5) = WBS6
            Report.Cells(ReportRow, 6) = Tabelle24.Range("I" & sourceRow).Value
            Report.Cells(ReportRow, 7) = Tabelle24.Range("K" & sourceRow).Value
            Report.Cells(ReportRow, 8) = Tabelle24.Range("L" & sourceRow).Value
            Report.Cells(ReportRow, 9) = Tabelle24.Range("N" & sourceRow).Value
            Report.Cells(ReportRow, 11) = 0
            
        ElseIf Tabelle24.Range("O" & sourceRow).Value > 0 And Tabelle24.Range("O" & sourceRow).Value <= Tabelle24.Range("P" & sourceRow).Value _
         And WBS = "A" Then
        
            ReportRow = Report.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            Report.Cells(ReportRow, 1) = WBS2
            Report.Cells(ReportRow, 2) = WBS3
            Report.Cells(ReportRow, 3) = WBS4
            Report.Cells(ReportRow, 4) = WBS5
            Report.Cells(ReportRow, 5) = WBS6
            Report.Cells(ReportRow, 6) = Tabelle24.Range("I" & sourceRow).Value
            Report.Cells(ReportRow, 7) = Tabelle24.Range("K" & sourceRow).Value
            Report.Cells(ReportRow, 8) = Tabelle24.Range("L" & sourceRow).Value
            
            If Tabelle24.Range("N" & sourceRow).Value < 0 Then
            
                Report.Cells(ReportRow, 9) = "Started"
                
            End If
            
            Report.Cells(ReportRow, 10) = Tabelle24.Range("O" & sourceRow).Value
            
            If Tabelle24.Range("N" & sourceRow).Value < 0 Then
            
                status = (Tabelle24.Range("N" & sourceRow).Value * -1) / Tabelle24.Range("P" & sourceRow).Value
            
                Report.Cells(ReportRow, 11) = status
                
            End If
            
        ElseIf Tabelle24.Range("O" & sourceRow).Value < 0 And WBS = "A" Then
        
            fnshRow = Finished.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            Finished.Cells(fnshRow, 1) = WBS2
            Finished.Cells(fnshRow, 2) = WBS3
            Finished.Cells(fnshRow, 3) = WBS4
            Finished.Cells(fnshRow, 4) = WBS5
            Finished.Cells(fnshRow, 5) = WBS6
            Finished.Cells(fnshRow, 6) = Tabelle24.Range("I" & sourceRow).Value
            Finished.Cells(fnshRow, 7) = Tabelle24.Range("K" & sourceRow).Value
            Finished.Cells(fnshRow, 8) = Tabelle24.Range("L" & sourceRow).Value
            Finished.Cells(fnshRow, 9) = "Finished"
            
        End If
         
Next sourceRow
        

End Sub


