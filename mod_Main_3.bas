Attribute VB_Name = "mod_Main_3"
Option Explicit

' Deklaration der globalen Variablen
Public ws               As Worksheet            ' Allgemeine Arbeitsblatt-Variable
Public wsSettings       As Worksheet            ' Arbeitsblatt für Daten
Public wsDaten          As Worksheet            ' Arbeitsblatt für Daten
Public wsHK             As Worksheet            ' Arbeitsblatt für HK
Public wsPE             As Worksheet            ' Arbeitsblatt für PE
Public wsVerListe       As Worksheet            ' Arbeitsblatt für VerListe
Public wsEBW            As Worksheet            ' Arbeitsblatt für EBW
Public wsVR             As Worksheet            ' Arbeitsblatt für Vertriebsreport

Public tbl              As ListObject           ' Allgemeine ListObject-Variable
Public loZielDaten      As ListObject           ' ListObject für Ziel
Public loQuelleDaten    As ListObject           ' ListObject für Quelle
Public loZiel           As ListObject           ' ListObject für Ziel
Public loQuelle         As ListObject           ' ListObject für Quelle
Public tblVR            As ListObject           ' ListObject für Vertriebsreport
Public tblHK            As ListObject           ' ListObject für Vertriebsreport

Public lastRow          As Long
Public spAnzahl         As Long                 ' Spaltenanzahl
Public zeAnzahl         As Long                 ' Zeilenanzahl

Public bolFlg           As Boolean              ' Allgemeiner Flag für Bedingungen
Public initializing     As Boolean              ' Flag für Initialisierung

Public cbMonth          As String               ' Ausgewählter Monat in der ComboBox
Public cbYear           As String               ' Ausgewähltes Jahr in der ComboBox
Public Datum1           As String               ' Formatierter Datumswert für Jahr und Monat

Public Sub FormatHardKopy()

    ' Setzen von Tabellenverweisen
    Set loZiel = wsHK.ListObjects("tbl_HK")

    wsHK.Activate
    With wsHK
        ' Autofit für alle Spalten
        .ListObjects("tbl_HK").Range.Columns.AutoFit
        ' Löschen der Spalte "Produktgruppenebene alternativ 1"
        .Range("tbl_HK[Produktgruppenebene alternativ 1]").Delete Shift:=xlToLeft
        ' Bereich Kunden
        .ListObjects("tbl_HK").ListColumns.Add Position:=2
        .ListObjects("tbl_HK").ListColumns.Add Position:=3
        .Range("A1").Value = "Kunden"
        .Range("B1").Value = "Kunden-Nr."
        .Range("C1").Value = "Kunde"
        ' Formeln für "Kunden-Nr." und "Kunde" einfügen
        .Range("tbl_HK[Kunden-Nr.]").formula = "=LEFT(RC[-1],SEARCH("" "",RC[-1])-1)"
        .Range("tbl_HK[Kunde]").formula = "=TRIM(MID(RC[-2],FIND("" "",RC[-2])+1,LEN(RC[-2])-FIND("" "",RC[-2])))"
        ' Werte aktualisieren

        ' Autofit für alle Spalten
        .ListObjects("tbl_HK").Range.Columns.AutoFit
        
        .Range("I1").Select
        ' Bereich Kunden
        .ListObjects("tbl_HK").ListColumns.Add Position:=5
        .Range("E1").Value = "Leander_Code"
        .Range("tbl_HK[Leander_Code]").formula = "=LEFT(RC[-1],SEARCH("" "",RC[-1])-1)"
        .Range("tbl_HK[[Kunden-Nr.]:[Leander_Code]]").Value = .Range("tbl_HK[[Kunden-Nr.]:[Leander_Code]]").Value
        Columns("D:D").Select
        Selection.Delete Shift:=xlToLeft
        
         .Range("I1").Select
        ' Bereich Produktgruppen PG_Ebene
        .ListObjects("tbl_HK").ListColumns.Add Position:=6
        .ListObjects("tbl_HK").ListColumns.Add Position:=7
        .Range("F1").Value = "PG_Ebene"
        ' Formel für "PG_Ebene" einfügen
        .Range("tbl_HK[PG_Ebene]").formula = "=LEFT(RC[-1],SEARCH("" "",RC[-1])-1)"
        Call SetTextFormat(wsHK, "F")
        ' Wert aktualisieren
        .Range("tbl_HK[PG_Ebene]").Value = .Range("tbl_HK[PG_Ebene]").Value
        ' Bereich Produktgruppen PG1
        .Range("G1").Value = "PG"
        ' Formel für "PG 1" einfügen
        .Range("tbl_HK[PG]").formula = "=TRIM(MID(RC[-2],FIND("" "",RC[-2])+1,LEN(RC[-2])-FIND("" "",RC[-2])))"
        ' Wert aktualisieren
        .Range("tbl_HK[PG]").Value = .Range("tbl_HK[PG]").Value
        .Range("I1").Select
    
        ' Bereich Produktgruppen PGA 1 & PGA 2
        .ListObjects("tbl_HK").ListColumns.Add Position:=9
        .ListObjects("tbl_HK").ListColumns.Add Position:=10
        .Range("I1").Value = "PGA_Nr"
        .Range("J1").Value = "PGA"
        ' Autofit für alle Spalten
        .ListObjects("tbl_HK").Range.Columns.AutoFit
        ' Formel für "PGA 1" einfügen
        .Range("tbl_HK[PGA_Nr]").formula = "=LEFT(RC[-1],FIND(""  "",RC[-1])-1)"
        Call SetTextFormat(wsHK, "I")
        ' Wert aktualisieren
        .Range("tbl_HK[PGA_Nr]").Value = .Range("tbl_HK[PGA_Nr]").Value
        ' Formel für "PGA 2" einfügen
        .Range("tbl_HK[PGA]").formula = "=REPLACE(MID(RC[-2],FIND("" "",RC[-2],FIND("" "",RC[-2])+1)+1,LEN(RC[-2])), 1, LEN(MID(RC[-2],FIND("" "",RC[-2],FIND("" "",RC[-2])+1)+1,LEN(RC[-2]))) - LEN(TRIM(MID(RC[-2],FIND("" "",RC[-2],FIND("" "",RC[-2])+1)+1,LEN(RC[-2])))), """")"
        ' Wert aktualisieren
        .Range("tbl_HK[PGA]").Value = .Range("tbl_HK[PGA]").Value

        ' Bereich Umsatz
        .Range("L1").Value = "Umsatz"
        .Range("M1").Value = "HK"
        .Range("N1").Value = "LAP_Lager"
        .Range("O1").Value = "WAP_Werk"
        
        Call SinEurFormat(wsHK, "L")
        Call SinEurFormat(wsHK, "M")
        Call SinEurFormat(wsHK, "N")
        Call SinEurFormat(wsHK, "O")
        ' Autofit für alle Spalten
        .ListObjects("tbl_HK").Range.Columns.AutoFit
    End With

' Löschen von Spalten A, D und F
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft
Columns("D:D").Select
Selection.Delete Shift:=xlToLeft
Columns("F:F").Select
Selection.Delete Shift:=xlToLeft

With wsHK
    ' Hinzufügen von Spalten "AD MA", "Art." und "PE Händler" zur Tabelle
    .ListObjects("tbl_HK").ListColumns.Add Position:=8
    .ListObjects("tbl_HK").ListColumns.Add Position:=9
    .ListObjects("tbl_HK").ListColumns.Add Position:=10
    .ListObjects("tbl_HK").ListColumns.Add Position:=11
    .ListObjects("tbl_HK").ListColumns.Add Position:=12
    .Range("H1").Value = "AD MA"
    .Range("I1").Value = "Art."
    .Range("J1").Value = "PE Händler"
    .Range("K1").Value = "Gebiet"
    .Range("L1").Value = "IC"
End With

' Verwendung von VLOOKUP-Formeln, um die Werte in den Spalten "PE Händler", "AD MA" und "Art." un IC zu füllen
Set loQuelle = wsPE.ListObjects("tbl_PE")
loZiel.ListColumns("PE Händler").DataBodyRange.formula = "=IFERROR(VLOOKUP([@[Kunden-Nr.]]," & "'" & wsPE.Name & "'" & "!" & loQuelle.Range.Address & ",2,FALSE),"""")"
Set loQuelle = wsVerListe.ListObjects("tbl_VerList")
loZiel.ListColumns("AD MA").DataBodyRange.formula = "=IFERROR(VLOOKUP([@[Kunden-Nr.]]," & "'" & wsVerListe.Name & "'" & "!" & loQuelle.Range.Address & ",4,FALSE),"""")"
Set loQuelle = wsEBW.ListObjects("tbl_EBW")
loZiel.ListColumns("Art.").DataBodyRange.formula = "=IFERROR(VLOOKUP([@[PGA_Nr]]," & "'" & wsEBW.Name & "'" & "!" & loQuelle.Range.Address & ",3,FALSE),"""")"
Set loQuelle = wsVerListe.ListObjects("tbl_VerList")
loZiel.ListColumns("Gebiet").DataBodyRange.formula = "=IFERROR(VLOOKUP([@[Kunden-Nr.]]," & "'" & wsVerListe.Name & "'" & "!" & loQuelle.Range.Address & ",6,FALSE),"""")"

' IC noch testen!
loZiel.ListColumns("IC").DataBodyRange.formula = _
"=IF(ISNUMBER(MATCH([@Kunde],tbl_IC[Auswahl IC],0)),""JA"","""")"


' Kopieren und Einfügen der Werte in den Spalten "AD MA" und "PE Händler"
Range("tbl_HK[[AD MA]:[IC]]").Select
Application.CutCopyMode = False
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Selection.NumberFormat = "@"

' Ersetzen von #NV-Fehlern durch leere Zellen
Selection.Replace What:="#NV", Replacement:="", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Application.CutCopyMode = False

' Anpassen der Zellenbreite für die gesamte Tabelle
Cells.Select
Cells.EntireColumn.AutoFit
' Löschen der Zeilen "spAnzahl - 1" und "spAnzahl"
Rows(spAnzahl & ":" & spAnzahl).Select
Selection.Delete Shift:=xlUp

' Zelle A1 auswählen
Range("A1").Select
Debug.Print "ENDE - HardKopy"
End Sub
' Prozedur zum Erstellen einer "HardKopy" des Datenblattes
Public Sub CreateVertriebsreport()
    ' Initialisiere Flag
    bolFlg = False

    ' Überprüfe, ob ein Arbeitsblatt mit dem Namen "HardKopy" bereits existiert
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Vertriebsreport" Then
            bolFlg = True
            Exit For
        End If
    Next ws

    ' Wenn kein Arbeitsblatt mit dem Namen "Vertriebsreport" gefunden wurde, erstelle ein neues
    If Not bolFlg Then
        ' Füge ein neues Arbeitsblatt nach dem aktiven Arbeitsblatt ein
        Sheets.Add After:=ActiveSheet
        ' Benenne das neue Arbeitsblatt "Vertriebsreport"
        ActiveSheet.Name = "Vertriebsreport"
        ' Setze Verweis auf das neu erstellte Arbeitsblatt "HardKopy"
        Set wsVR = ThisWorkbook.Sheets("Vertriebsreport")
    End If

    ' Kopiere den Bereich aus dem Datenblatt in das "Vertriebsreport"-Arbeitsblatt
    spAnzahl = wsHK.Cells(wsHK.Rows.Count, 1).End(xlUp).Row
    wsHK.Range("A1:J" & spAnzahl).Copy
    wsVR.Range("A1").PasteSpecial xlPasteValues
    ' Erstelle eine Tabelle für den kopierten Bereich im "HardKopy"-Arbeitsblatt
    Set tbl = wsVR.ListObjects.Add(xlSrcRange, wsVR.Range("A1:P" & spAnzahl), , xlYes)
    ' Benenne die Tabelle "tbl_HK"
    tbl.Name = "tbl_VR"
    
    With wsVR
        .Range("A1").Value = "Kunden_Nr"
        .Range("B1").Value = "Kunde"
        .Range("C1").Value = "Leander_Code"
        .Range("D1").Value = "PG_Ebene"
        .Range("E1").Value = "PG"
        .Range("F1").Value = "PGA_Nr"
        .Range("G1").Value = "PGA"
        .Range("H1").Value = "Monat"
        .Range("I1").Value = "Umsatz"
        .Range("J1").Value = "HK"
        .Range("K1").Value = "LAP_Lager"
        .Range("L1").Value = "WAP_Werk"
        .Range("M1").Value = "Kosten_DB1"
        .Range("N1").Value = "Marge_DB1"
        .Range("O1").Value = "Marge_DB1_Prozent"
        .Range("P1").Value = "Zuschlaege_DB3"
        .Range("Q1").Value = "Kosten_DB3"
        .Range("R1").Value = "Marge_DB3"
        .Range("S1").Value = "Marge_DB3_Prozent"
        .Range("T1").Value = "AD_MA"
        .Range("U1").Value = "PE_Haendler"
        .Range("V1").Value = "EinbauWerkZ"
        .Range("W1").Value = "Gebiet"
        .Range("X1").Value = "IC_Gesellschaft"
    End With

    ' Finde die letzte Zeile in der "HardKopy" Tabelle
    lastRow = wsHK.Cells(wsHK.Rows.Count, 1).End(xlUp).Row

    ' Kopiere die Daten aus der "HardKopy" Tabelle in die entsprechenden Spalten der "Vertriebsreport" Tabelle
    wsHK.Range("A2:A" & lastRow).Copy wsVR.Range("A2") 'Kunden-Nr.
    wsHK.Range("B2:B" & lastRow).Copy wsVR.Range("B2") 'Kunde
    wsHK.Range("C2:C" & lastRow).Copy wsVR.Range("C2") 'Leander_Code
    wsHK.Range("D2:D" & lastRow).Copy wsVR.Range("D2") 'PG_Ebene
    wsHK.Range("E2:E" & lastRow).Copy wsVR.Range("E2") 'PG
    wsHK.Range("F2:F" & lastRow).Copy wsVR.Range("F2") 'PGA_Nr
    wsHK.Range("G2:G" & lastRow).Copy wsVR.Range("G2") 'PGA
    wsHK.Range("M2:M" & lastRow).Copy wsVR.Range("H2") 'Monat
    wsHK.Range("N2:N" & lastRow).Copy wsVR.Range("I2") 'Umsatz

    'Masure
    wsHK.Range("O2:O" & lastRow).Copy wsVR.Range("J2") 'HK
    wsHK.Range("P2:P" & lastRow).Copy wsVR.Range("K2") 'LAP_Lager
    wsHK.Range("Q2:Q" & lastRow).Copy wsVR.Range("L2") 'WAP_Werk

    'Verknüpfungen
    wsHK.Range("H2:H" & lastRow).Copy wsVR.Range("T2") 'AD_MA
    wsHK.Range("J2:J" & lastRow).Copy wsVR.Range("U2") 'PE_Haendler
    wsHK.Range("I2:I" & lastRow).Copy wsVR.Range("V2") 'EinbauWerkZ
    wsHK.Range("K2:K" & lastRow).Copy wsVR.Range("W2") 'Gebiet
    wsHK.Range("L2:L" & lastRow).Copy wsVR.Range("X2") 'IC_Gesellschaft

    wsVR.ListObjects("tbl_VR").Range.Columns.AutoFit
    
    Debug.Print "ENDE - CreateVertriebsreport"
    
End Sub

Public Sub TransferDataToVertriebsreport()

    Dim lastRow As Long
    Dim formulaText As String
    Dim formula As String
    Dim cell As Range
    Dim Row As Long
    Dim pgEbene As Variant   ' Declare pgEbene as Variant
    
    Set tblVR = wsVR.ListObjects("tbl_VR")
    
    
    
    ' Finde die letzte Zeile in der "HardKopy" Tabelle
    lastRow = wsHK.Cells(wsHK.Rows.Count, 1).End(xlUp).Row
    
    ' Setze die Formel direkt in die erste Zelle der Spalte "Kosten_DB1" 'wsVR.Range("M2")'
    With tblVR.ListColumns("Kosten_DB1").DataBodyRange
        .formula = "=IFERROR(([@HK]*0.0674)+[@[LAP_Lager]], """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
     End With
     
    ' Marge_DB1
    With tblVR.ListColumns("Marge_DB1").DataBodyRange
        .formula = "=IFERROR([@Umsatz]-[@[Kosten_DB1]], """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
    End With
    
    ' Marge_DB1_Prozent
    With tblVR.ListColumns("Marge_DB1_Prozent").DataBodyRange
        .formula = "=IFERROR(([@[Marge_DB1]]/[@Umsatz]), """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
        .NumberFormat = "0.00%"
    End With
    
'    ' Die Formeln in der Spalte "Zuschlaege_DB3" in Werte umwandeln
'    With tblVR.ListColumns("Zuschlaege_DB3").DataBodyRange
'
'        .Value = .Value
'    End With

    Dim formulaRange As Range
    Dim formulaCell As Range
    Set formulaRange = wsSettings.Range("A10:A18") ' Angepasst an den Bereich, in dem sich die Formeln befinden


    
    For Each cell In tblVR.ListColumns("PG_Ebene").DataBodyRange
        Row = cell.Row - tblVR.HeaderRowRange.Row
        pgEbene = cell.Value
    
        For Each formulaCell In formulaRange
            If formulaCell.Value = pgEbene Then
                formula = wsSettings.Cells(formulaCell.Row, 2).Value ' Die Formel befindet sich in der zweiten Spalte des formulaRange-Bereichs
                formula = Application.WorksheetFunction.Substitute(formula, "HK", tblVR.DataBodyRange.Cells(Row, tblVR.ListColumns("HK").Index).Address(, , xlR1C1))
                formula = Replace(formula, ",", ".")
                tblVR.DataBodyRange.Cells(Row, tblVR.ListColumns("Zuschlaege_DB3").Index).formula = "=IFERROR(" & formula & ", """")"
                tblVR.DataBodyRange.Cells(Row, tblVR.ListColumns("Zuschlaege_DB3").Index).NumberFormat = "0.000000" ' 6 Dezimalstellen
                tblVR.DataBodyRange.Cells(Row, tblVR.ListColumns("Zuschlaege_DB3").Index).Value = tblVR.DataBodyRange.Cells(Row, tblVR.ListColumns("Zuschlaege_DB3").Index).Value
                Exit For
            End If
        Next formulaCell
    Next cell
     
    ' Kosten_DB3
    With tblVR.ListColumns("Kosten_DB3").DataBodyRange
        .formula = "=[@[Kosten_DB1]]+[@[LAP_Lager]]"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
    End With

    ' Kosten_DB3
    With tblVR.ListColumns("Kosten_DB3").DataBodyRange
        .formula = "=IFERROR([@[Kosten_DB1]]+[@[Zuschlaege_DB3]], """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
    End With

    ' Marge_DB3
    With tblVR.ListColumns("Marge_DB3").DataBodyRange
        .formula = "=IFERROR([@Umsatz]-[@[Kosten_DB3]], """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
    End With

    ' Marge_DB3_Prozent
    With tblVR.ListColumns("Marge_DB3_Prozent").DataBodyRange
        .formula = "=IFERROR(([@[Marge_DB3]]/[@Umsatz]), """")"
        .NumberFormat = "0.000000" ' 6 Dezimalstellen
        .Value = .Value
        .NumberFormat = "0.00%"
    End With
    
    Call SinEurFormat(wsVR, "J") ', "K", "L", "M", "N", "P", "Q", "R")
    Call SinEurFormat(wsVR, "K")
    Call SinEurFormat(wsVR, "L")
    Call SinEurFormat(wsVR, "M")
    Call SinEurFormat(wsVR, "N")
    Call SetPercentFormat(wsVR, "O")
    Call SinEurFormat(wsVR, "P")
    Call SinEurFormat(wsVR, "Q")
    Call SinEurFormat(wsVR, "R")
    Call SetPercentFormat(wsVR, "S")

    wsVR.ListObjects("tbl_VR").Range.Columns.AutoFit
    
    Debug.Print "ENDE - TransferDataToVertriebsreport"
End Sub


