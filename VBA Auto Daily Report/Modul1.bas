Attribute VB_Name = "Modul1"
Option Explicit

Dim n As Integer        'n  - Anzahl der Zeilen
Dim s As Integer        's  - Anzahl der Spalten
Dim iRow As Variant     'iRow  - Zeilennummer
Dim zz As Integer       'zz - Textlänge
Dim x As Integer        'x  - SchleifenZähler
Dim strText As String
Dim lz As Integer       'lz - Leerzeichen
Dim Level As Integer    'Level = Ersatzz für Leerzeichen
Dim MaxLevel As Integer 'max. Anzahl der Ebenen
Dim Farbe(1 To 15) As Long
Dim Schriftfarbe(1 To 15) As Long
Dim Schriftgroesse(1 To 15) As Long






Sub Tabelle_Formatieren()
Attribute Tabelle_Formatieren.VB_ProcData.VB_Invoke_Func = "T\n14"


MaxLevel = 0
n = 0
s = 0
iRow = 0


'Farbwerte als RGB aus Primavera für Farbstruktur
Farbe(1) = RGB(255, 51, 51)   'Red Colour
Farbe(2) = RGB(0, 0, 204)     'Blue Colour
Farbe(3) = RGB(102, 255, 0)   'Green colour
Farbe(4) = RGB(255, 255, 51)  'Yellow Colour
Farbe(5) = RGB(153, 153, 153) 'Grey Colour
Farbe(6) = RGB(204, 204, 204) 'Light Grey colour
Farbe(7) = RGB(255, 204, 255) 'LightPink Colour
Farbe(8) = RGB(255, 255, 153) 'Light Yellow
Farbe(9) = RGB(51, 51, 51)    'Light Black
Farbe(10) = RGB(0, 0, 0)
Farbe(11) = RGB(0, 0, 0)
Farbe(12) = RGB(0, 0, 0)
Farbe(13) = RGB(0, 0, 0)
Farbe(14) = RGB(0, 0, 0)
Farbe(15) = RGB(0, 0, 0)

'Schriftfarben der Ebenen
Schriftfarbe(1) = RGB(255, 255, 255)
Schriftfarbe(2) = RGB(255, 255, 255)
Schriftfarbe(3) = RGB(0, 0, 0)
Schriftfarbe(4) = RGB(0, 0, 0)
Schriftfarbe(5) = RGB(255, 255, 255)
Schriftfarbe(6) = RGB(0, 0, 0)
Schriftfarbe(7) = RGB(0, 0, 0)
Schriftfarbe(8) = RGB(0, 0, 0)
Schriftfarbe(9) = RGB(255, 255, 255)
Schriftfarbe(10) = RGB(0, 0, 0)
Schriftfarbe(11) = RGB(0, 0, 0)
Schriftfarbe(12) = RGB(0, 0, 0)
Schriftfarbe(13) = RGB(0, 0, 0)
Schriftfarbe(14) = RGB(0, 0, 0)
Schriftfarbe(15) = RGB(0, 0, 0)

'Schriftgroesse der Ebenen
Schriftgroesse(1) = 14
Schriftgroesse(2) = 13
Schriftgroesse(3) = 12
Schriftgroesse(4) = 12
Schriftgroesse(5) = 11
Schriftgroesse(6) = 10
Schriftgroesse(7) = 10
Schriftgroesse(8) = 9
Schriftgroesse(9) = 9
Schriftgroesse(10) = 6
Schriftgroesse(11) = 6
Schriftgroesse(12) = 6
Schriftgroesse(13) = 6
Schriftgroesse(14) = 6
Schriftgroesse(15) = 6


'Copyright: Stefan Mollenkopf 2008
'Modified: Hans-Dieter Rapp 12/2008
'Modified: Jörg Weber 2008-11
On Error GoTo fehler
'Alles löschen und Daten einfügen
    Columns("A:IU").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    MsgBox "Data Copy & Press OK :) :)", , "Meldung"
 Application.DisplayAlerts = False
    ActiveSheet.Paste
Application.DisplayAlerts = True
Application.ScreenUpdating = False
    Sheets("Tabelle1").Select

'Zeilenzähler
    For Each iRow In ActiveCell.CurrentRegion.Rows
        n = n + 1
    Next

'Spaltenzähler
    For Each iRow In ActiveCell.CurrentRegion.Columns
        s = s + 1
    Next

'Tabelle verschieben
    Range(Cells(1, 1), Cells(n, s)).Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    
    

    For iRow = 4 To n + 2
        If Cells(iRow, 2) = "" Then
            strText = Cells(iRow, 1).Value
            lz = Len(strText) - Len(LTrim(strText))
            If lz / 2 > MaxLevel Then MaxLevel = lz / 2
        End If
    Next
        
   MaxLevel = MaxLevel + 1
   MsgBox "Maxlevel: " & MaxLevel, , "Meldung"
     
   

'Spalten vor die Tabelle einfügen und den Hintergrund festlegen
    For x = 1 To MaxLevel
        Range("A1").Select
        Selection.EntireColumn.Insert
        Columns("A:A").ColumnWidth = 1
    Next x
        


'Schleifendurchlauf
    For iRow = 4 To n + 2
       
        lz = 0
    
        If Cells(iRow, 2 + MaxLevel) <> "" Then
    
' Keine_Ebene - sondern Vorgang
            Range(Cells(iRow, 1 + MaxLevel), Cells(iRow, s + MaxLevel)).Select
            With Selection
                .Font.Size = 9
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
            End With
        
            Range(Cells(iRow, 2 + MaxLevel), Cells(iRow, s + MaxLevel)).WrapText = True
          
            Range(Cells(iRow, 1 + MaxLevel), Cells(iRow, s + MaxLevel)).Select
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
        Else
    
            
            strText = Cells(iRow, 1 + MaxLevel).Value
            For zz = 1 To Len(strText)
                If Mid(strText, zz, 1) = " " Then
                    lz = lz + 1
                End If
                If Mid(strText, zz + 1, 1) <> " " Then Exit For
            Next zz
            lz = Len(strText) - Len(LTrim(strText))

        
'Ebene 1 bis Maxlevel Formatierung

            For x = 1 To MaxLevel
                Level = x * 2 - 2
                If lz = Level Then
                    Range(Cells(iRow, 1 + x), Cells(iRow, s + MaxLevel)).Select
                    Selection.Interior.Color = Farbe(x)
                    Selection.Interior.Pattern = xlSolid
                    Selection.Font.Name = "Arial"
                    Selection.Font.FontStyle = "Fett"
                    Selection.Font.Size = Schriftgroesse(x)
                    Selection.Font.Color = Schriftfarbe(x)
                    
                    
                    Range(Cells(iRow, 1 + lz / 2), Cells(n + 2, 1 + lz / 2)).Select
                    Selection.Interior.Color = Farbe(x)
                    Selection.Interior.Pattern = xlSolid
                    

                    
                    If (2 + lz / 2) < (MaxLevel) Then
                        Range(Cells(iRow + 1, 2 + lz / 2), Cells(n + 2, MaxLevel)).Select
                        Selection.Interior.Pattern = xlNone
                        'Selection.Interior.TintAndShade = 0
                        'Selection.Interior.PatternTintAndShade = 0
                    End If


'                    Range(Cells(iRow, 1 + x), Cells(iRow, MaxLevel - 1)).Select
'                    Selection.HorizontalAlignment = xlLeft
'                    Selection.VerticalAlignment = xlTop
'                    Selection.MergeCells = True
'                    Selection.WrapText = True

                End If
            Next
        End If
    Next


'Breite und Rahmen festlegen

Columns(MaxLevel + 1).ColumnWidth = 16     'Spalte Vorgangs-ID

Columns(MaxLevel + 2).ColumnWidth = 50       'Spalte Vorgansname
Columns(MaxLevel + 2 + 53).ColumnWidth = 50     'Kontroll-Spalte Vorgansname

Columns(MaxLevel + 4).Select                 'Spalte Responsibility
    With Selection
        .ColumnWidth = 14
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

Columns(MaxLevel + 5).Select                 'Spalte AIC
    With Selection
        .ColumnWidth = 14
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With


Columns(MaxLevel + 6).Select                 'Spalte Start
    With Selection
        .ColumnWidth = 14
        .HorizontalAlignment = xlCenter
    End With
        
        
Columns(MaxLevel + 7).Select                 'Spalte Finish
    With Selection
        .ColumnWidth = 14
        .HorizontalAlignment = xlCenter
    End With


Columns(MaxLevel + 8).Select                 'Spalte Expected Start
    With Selection
            .ColumnWidth = 10
            .HorizontalAlignment = xlCenter
            .WrapText = True
    End With
    
    
Columns(MaxLevel + 9).Select                 'Spalte Expected Finish
    With Selection
            .ColumnWidth = 10
            .HorizontalAlignment = xlCenter
            .WrapText = True
    End With


Columns(MaxLevel + 10).ColumnWidth = 20        'Spalte Comments

Range("J:R").Font.Size = 9

Range(Cells(3, 13), Cells(n + 2, 15)).Select
    With Selection
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .NumberFormat = "dd/mm/yy;@"
        
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
       
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
    
    End With


'Überschrift Formatierung
        Range("A3:J3").Select
            With Selection
                .MergeCells = False
            End With
        Range(Cells(3, 1), Cells(3, s + 9)).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .ColorIndex = xlAutomatic
            End With
            With Selection.Font
                .Size = 9
                .FontStyle = "Fett"
            End With
            
            
'Überschriftzeile ausrichten
            'Range("L3").FormulaR1C1 = "Responsibilities"   'Umbenennung Überschrift
            'Range("M3").FormulaR1C1 = "AIC"                'Umbenennung Überschrift
            'Range(Cells(3, 1), Cells(3, s + 9)).Select
             '   With Selection
              '      .WrapText = True
               '     .HorizontalAlignment = xlCenter
                '    .VerticalAlignment = xlCenter
                'End With


'Rahmen für ganze Tabelle
Range(Cells(3, 1), Cells(n + 2, s + 9)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
    End With

'Statusdatum
Cells(1, s + MaxLevel - 1).Select
    With Selection
        .FormulaR1C1 = "Status:"
        .HorizontalAlignment = xlRight
    End With
    
Cells(1, s + MaxLevel + 1).Select
    With Selection
        .FormulaR1C1 = "=today()"
        .HorizontalAlignment = xlLeft
    End With
Cells(1, s + MaxLevel).Value = Cells(1, s + MaxLevel + 1).Value
Cells(1, s + MaxLevel + 1).Value = ""

'Überschrift
Range("A1").Select
    With Selection
        .FormulaR1C1 = "       "
        .HorizontalAlignment = xlLeft
    End With
    With Selection.Font
        .FontStyle = "Fett"
        .Size = 12
    End With

'Höhe anpassen
ActiveSheet.Rows.AutoFit

'Hilfe-Spalten löschen
'    Range(Cells(1, 32), Cells(n + 2, 33)).ClearContents




fehler:
Application.ScreenUpdating = True

End Sub


Sub P6DatetoExcel()
Attribute P6DatetoExcel.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' P6DatetoExcel Makro
' Makro am 04.05.2010 von Weber aufgezeichnet
'
' Tastenkombination: Strg+Umschalt+P
'
Dim a As Variant
Dim dr As Integer

Application.ScreenUpdating = False


dr = Selection.Column
Range(Cells(2, dr), Cells(2, dr)).Select
n = 0
For Each iRow In ActiveCell.CurrentRegion.Rows
        n = n + 1
    Next
    n = n + 1
    

Range(Cells(1, dr), Cells(n, dr)).Select
On Error GoTo fehler
    Selection.Replace What:="jan", Replacement:="01", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="feb", Replacement:="02", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="mar", Replacement:="03", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="mrz", Replacement:="03", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="apr", Replacement:="04", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="may", Replacement:="05", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="mai", Replacement:="05", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="jun", Replacement:="06", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="jul", Replacement:="07", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="aug", Replacement:="08", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="sep", Replacement:="09", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="oct", Replacement:="10", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="okt", Replacement:="10", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="nov", Replacement:="11", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="dez", Replacement:="12", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="dec", Replacement:="12", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="~*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.WrapText = False
    Application.ScreenUpdating = True
    
    
    'Datumwerte auf das gleiche Format bringen und Referen kopieren
    
    Range(Cells(1, dr + 53), Cells(n, dr + 53)).Select
    Selection.FormulaR1C1 = "=IFERROR(DATEVALUE(RC[-53]),IF(RC[-53]=0,"""",RC[-53]))"
    Selection.Copy
    Range(Cells(1, dr), Cells(n, dr)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "m/d/yyyy"
    Selection.Copy
    Range(Cells(1, dr + 53), Cells(n, dr + 53)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "m/d/yyyy"
    
    Cells(1, dr).EntireColumn.AutoFit
     Cells(1, dr).Select
     Selection.EntireColumn.AutoFit
    Dim rngCell As Range
    For Each rngCell In Selection
      rngCell.FormulaLocal = rngCell.FormulaLocal
    Next rngCell
        
fehler:


End Sub

Sub Überwachung()
Attribute Überwachung.VB_ProcData.VB_Invoke_Func = "U\n14"
Application.ScreenUpdating = False

'Überwachung
    Range(Cells(1, 1), Cells(65535, 53)).Select
    Selection.Copy
    
    Range("bb1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.EntireColumn.Hidden = True
    Range(Cells(1, 1), Cells(65535, 53)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=bb1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    ActiveWindow.LargeScroll ToRight:=1
    ActiveWindow.LargeScroll ToRight:=-1
    Range("A1").Select
Application.ScreenUpdating = True

End Sub


Sub AdditionalColumns()
Attribute AdditionalColumns.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' AdditionalColumns Makro
'
' Tastenkombination: Strg+Umschalt+C
'
Dim Reihe As Integer

Application.ScreenUpdating = False

    Reihe = ActiveCell.Column
    Cells(3, Reihe).Select
    ActiveCell.FormulaR1C1 = "Start" & Chr(10) & "Actual/New"
    With ActiveCell.Characters(Start:=1, Length:=16).Font
        .Name = "Arial"
        .FontStyle = "Fett"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Cells(3, Reihe + 1).Select
    
    ActiveCell.FormulaR1C1 = "Finish" & Chr(10) & "Actual/New"
    With ActiveCell.Characters(Start:=1, Length:=17).Font
        .Name = "Arial"
        .FontStyle = "Fett"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Cells(3, Reihe + 2).Select
    ActiveCell.FormulaR1C1 = "Duration" & Chr(10) & "Working days"
    With ActiveCell.Characters(Start:=1, Length:=21).Font
        .Name = "Arial"
        .FontStyle = "Fett"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Columns(Reihe - 1).Select
    
    Selection.Copy
   
    Range(Columns(Reihe), Columns(Reihe + 2)).Select
   
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range(Cells(1, Reihe - 2), Cells(1, Reihe - 1)).Select
    Selection.Cut
    Cells(1, Reihe + 1).Select
    ActiveSheet.Paste
    Application.ScreenUpdating = True
    
End Sub


