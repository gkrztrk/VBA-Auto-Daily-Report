Attribute VB_Name = "MailSender"
Sub Send_Mail()
Dim mailApp As Outlook.Application
Set mailApp = New Outlook.Application
Dim mailItem As Outlook.mailItem
Set mailItem = mailApp.CreateItem(olMailItem)
Dim RepRange As Range
Dim lr As Integer

Report.Activate

lr = Report.Cells(Rows.Count, 1).End(xlUp).Row

Set RepRange = Report.Range("A1:I" & lr)

RepRange.AutoFilter field:=9, Criteria1:="<11", Operator:=xlFilterValues, VisibleDropDown:=True

Dim str1, str2 As String

str1 = "<BODY style= font-size:15pt;font-family:Calibri>" & _
"Dear Gents, <br><br>Please kindly find attached report regarding upcoming activities.<br>"

str2 = "<br>Regards,<br>Goker."

On Error Resume Next

    With mailItem
        .To = "nezih.sengezer@nesma.com; ahmed.rezek@nesma.com; imran.afzal@nesma.com; ahmed.alshaykh@nesma.com; mehmet.kabatas@nesma.com"
        .CC = "yigit.yazici@nesma.com"
        .Subject = "Roads & Paving Schedule Reminder"
        .Display
        .HTMLBody = str1 & RangetoHTML(RepRange) & str2 & .HTMLBody
        End With
        ThisWorkbook.Save
        Report.ShowAllData
        mailItem.Attachments.Add ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    On Error GoTo 0
    
mailItem.Send


Set mailItem = Nothing
Set mailApp = Nothing
End Sub



Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    'wss.Range("B:B").SpecialCells(xlCellTypeVisible).Copy
    'wsd.Range("A1").PasteSpecial xlPasteValues
    rng.SpecialCells(xlCellTypeVisible).Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        '.Cells(1).PasteSpecial xlPasteAllUsingSourceTheme
        '.Cells(1).PasteSpecial xlPasteAllMergingConditionalFormats
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        '.Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

