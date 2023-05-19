VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Basis_3 
   Caption         =   "Basis1"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   OleObjectBlob   =   "uf_Basis_3.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "uf_Basis_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bt_Verbinden_Click()
    ConnectAndCreatePivotTable
    UpdateButtonStatus ' Aktualisieren des Button-Status
End Sub

' Klicken auf Schaltfl�che "Berechnen"
Private Sub bt_Berechnen_HK_Click()
    'OptimizeVBA False
    DeleteOldSheetHK
    'ClearTableContents "HardKopy", "tblHK" ' Leere die Tabelle "tblHK" im Arbeitsblatt "HardKopy"
    CreateHardKopy
    FormatHardKopy
    'UpdateButtonStatus
    'OptimizeVBA True
End Sub
' Dieser Sub l�scht ein Arbeitsblatt namens "HardKopy", falls es vorhanden ist.

Public Sub DeleteOldSheetHK()
    Dim blatt As Worksheet
    ' Durchlaufen aller Arbeitsbl�tter
    
    For Each blatt In Sheets
        ' Pr�fen, ob das Arbeitsblatt "HardKopy" hei�t
        If blatt.Name = "HardKopy" Then
            ' "HardKopy" ausw�hlen
            Sheets("HardKopy").Select
            ' Das ausgew�hlte Arbeitsblatt l�schen
            ActiveWindow.SelectedSheets.Delete
        End If
    Next blatt
End Sub

' Prozedur zum Erstellen einer "HardKopy" des Datenblattes
Public Sub CreateHardKopy()
    ' Initialisiere Flag
    bolFlg = False
    
    ' Setze Verweis auf das Datenblatt
    Set wsDaten = ThisWorkbook.Sheets("Daten")
    
    ' �berpr�fe, ob ein Arbeitsblatt mit dem Namen "HardKopy" bereits existiert
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "HardKopy" Then
            bolFlg = True
            ' Setze Verweis auf das vorhandene Arbeitsblatt "HardKopy"
            Set wsHK = ThisWorkbook.Sheets("HardKopy")
            Exit For
        End If
    Next ws

    ' Wenn kein Arbeitsblatt mit dem Namen "HardKopy" gefunden wurde, erstelle ein neues
    If Not bolFlg Then
        ' F�ge ein neues Arbeitsblatt nach dem aktiven Arbeitsblatt ein
        Sheets.Add After:=ActiveSheet
        ' Benenne das neue Arbeitsblatt "HardKopy"
        ActiveSheet.Name = "HardKopy"
        ' Setze Verweis auf das neu erstellte Arbeitsblatt "HardKopy"
    End If
    Set wsHK = ThisWorkbook.Sheets("HardKopy")
    
    ' Kopiere den Bereich aus dem Datenblatt in das "HardKopy"-Arbeitsblatt
    spAnzahl = wsDaten.Cells(wsDaten.Rows.Count, 1).End(xlUp).Row
    wsDaten.Range("A1:J" & spAnzahl).Copy
    wsHK.Range("A1").PasteSpecial xlPasteValues
    ' Erstelle eine Tabelle f�r den kopierten Bereich im "HardKopy"-Arbeitsblatt
    Set tbl = wsHK.ListObjects.Add(xlSrcRange, wsHK.Range("A1:J" & spAnzahl), , xlYes)
    ' Benenne die Tabelle "tbl_1"
    tbl.Name = "tbl_HK"
    
End Sub
' Start und Initializierung
Private Sub UserForm_Initialize()
    Dim iYear As Integer
    Dim iMonth As Integer
    Dim currentDate As Date
    Dim previousMonthDate As Date
    
    Debug.Print "Start und Initializierung"
    
    ' Pr�fen und erstellen der Arbeitsbl�tter
    If Not CheckAndCreateSheets() Then
        MsgBox "Fehler beim Erstellen der Arbeitsbl�tter.", vbCritical, "Fehler"
        Exit Sub
    End If

    'Zuweisen der Arbeitsbl�tter zu den Variablen nachdem sie erstellt wurden
    
    Debug.Print "Zuweisen der Arbeitsbl�tter zu den Variablen nachdem sie erstellt wurden"
    
    Set wsHK = ThisWorkbook.Worksheets("HardKopy")
    Set wsDaten = ThisWorkbook.Worksheets("Daten")
    Set wsVR = ThisWorkbook.Worksheets("Vertriebsreport")
    Set wsEBW = ThisWorkbook.Worksheets("EinBauWerkz")
    Set wsPE = ThisWorkbook.Worksheets("PE_Haendler")
    Set wsVerListe = ThisWorkbook.Worksheets("VerListe")
    Set wsSettings = ThisWorkbook.Worksheets("Settings")

    
    ' Pr�fen und erstellen der ListObjects
    Debug.Print "Pr�fen und erstellen der ListObjects"
    If Not CheckAndCreateListObjects() Then
        MsgBox "Fehler beim Erstellen der ListObjects.", vbCritical, "Fehler"
        Exit Sub
    End If

    ' Deklaration von Variablen f�r Jahr und Monat
    Debug.Print "Deklaration von Variablen f�r Jahr und Monat"

    ' ListPivotFieldNames
    
    ' Initialisierungsstatus auf True setzen
    Debug.Print "Initialisierungsstatus auf True setzen"
    initializing = True
    
    cboYear.Value = Year(Date)
    cboMonth.Value = MonthName(1)
    
  
    ' Initialisierung der ComboBox-Variablen
    Debug.Print "Initialisierung der ComboBox-Variablen"
    cbYear = cboYear.Value
    cbMonth = cboMonth.Value
    
    ' Deaktivieren der Schaltfl�chen beim Start des UserForms
    Debug.Print "Deaktivieren der Schaltfl�chen beim Start des UserForms"
    bt_Verbinden.Enabled = False
    bt_Pvfuellen.Enabled = False
    bt_Berechnen_HK.Enabled = False


    ' F�llen der cboYear ComboBox mit Jahren
    Debug.Print "F�llen der cboYear ComboBox mit Jahren"
    For iYear = Year(Date) - 3 To Year(Date) + 2
        cboYear.AddItem iYear
    Next iYear

    ' F�llen der cboMonth ComboBox mit Monatsnamen
    Debug.Print "F�llen der cboMonth ComboBox mit Monatsnamen"
    For iMonth = 1 To 12
        cboMonth.AddItem MonthName(iMonth)
    Next iMonth

    ' Standardauswahl f�r Jahr und Monat in den ComboBoxen
    Debug.Print "Standardauswahl f�r Jahr und Monat in den ComboBoxen"
    currentDate = Date
    previousMonthDate = DateAdd("m", -1, currentDate)

    cboYear.Value = Year(previousMonthDate)
    cboMonth.Value = MonthName(Month(previousMonthDate))

    
    ' Aufruf der UpdateButtonStatus-Prozedur, um die Schaltfl�chen-Status zu aktualisieren
    Debug.Print "Aufruf der UpdateButtonStatus-Prozedur, um die Schaltfl�chen-Status zu aktualisieren"
    UpdateDatum1
    UpdateButtonStatus
    
    ' Initialisierungsstatus auf False setzen
    Debug.Print "Initialisierungsstatus auf False setzen"
    initializing = False
    uf_Basis_3.Show vbModal
    
    Debug.Print "Start und Initializierung - beendet"
End Sub
Private Sub bt_Pvfuellen_Click()
    
    Dim yearVal As Integer
    Dim monthVal As Integer
    Dim Jahr As String
    Dim Monat As Integer
    Dim i As Integer
    
    Dim newMeasure As String
    Dim measureHerstellkosten As String
    Dim measureLAP As String
    Dim measureWAP As String
    
    ' Jahr und Monat aus den Comboboxen abrufen
    Debug.Print "Jahr und Monat aus den Comboboxen abrufen"
    yearVal = cboYear.Value
    monthVal = cboMonth.ListIndex + 1

    ' �berpr�fen, ob die Pivot-Tabelle bereits mit den ausgew�hlten Werten gef�llt ist
    Debug.Print "�berpr�fen, ob die Pivot-Tabelle bereits mit den ausgew�hlten Werten gef�llt ist"
    If IsPivotTableFilled(wsDaten, yearVal, monthVal) Then
        ' Wenn die Pivot-Tabelle bereits gef�llt ist, aktualisieren Sie einfach den Monatsfilter
        Debug.Print "Wenn die Pivot-Tabelle bereits gef�llt ist, aktualisieren Sie einfach den Monatsfilter"
        'UpdateMonthFilter yearVal, monthVal
    Else
        ' Rechnungswert entvernen um doppelte ausf�hrung zu vermeiden.
        Debug.Print "Rechnungswert entvernen um doppelte ausf�hrung zu vermeiden."
        ActiveSheet.PivotTables("pv_Daten").CubeFields( _
        "[Measures].[Rechnungswert (bereinigt)]").Orientation = xlHidden
        
        ' PivotTable-Einstellungen anpassen
        Debug.Print "PivotTable-Einstellungen anpassen"
        With wsDaten.PivotTables("pv_Daten")
            .TotalsAnnotation = False
            .VisualTotals = True
            .AllowMultipleFilters = True
            .VisualTotalsForSets = True
        End With
        
        wsDaten.PivotTables("pv_Daten").CubeFields("[ZEIT].[Monat]").Orientation = xlRowField
        
        wsDaten.PivotTables("pv_Daten").PivotFields("[ZEIT].[Monat].[Monat]").VisibleItemsList = Array("[ZEIT].[Monat].&[" & Datum1 & "]")
        
        ' Produktgruppe Alternativ hinzuf�gen und filtern
        Debug.Print "Produktgruppe Alternativ hinzuf�gen und filtern"
        wsDaten.PivotTables("pv_Daten").CubeFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ]").CreatePivotFields
        wsDaten.PivotTables("pv_Daten").PivotFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].[Produktgruppenebene alternativ 1]").VisibleItemsList = Array("")
        wsDaten.PivotTables("pv_Daten").PivotFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].[Produktgruppenebene alternativ 2]").VisibleItemsList = Array( _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140345]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140346]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140347]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140348]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140349]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140350]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140351]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140352]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140379]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140380]", _
            "[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000140381]")
            
        wsDaten.PivotTables("pv_Daten").PivotFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].[Produktgruppenebene alternativ 3]").VisibleItemsList = Array("")
              
        wsDaten.PivotTables("pv_Daten").CubeFields("[PRODUKT ALTERNATIV].[Produkt Alternativ]").Orientation = xlRowField
        
        ' KUNDE als erstes Zeilenfeld setzen und Position festlegen
        Debug.Print "KUNDE als erstes Zeilenfeld setzen und Position festlegen"
        With wsDaten.PivotTables("pv_Daten").CubeFields("[KUNDE].[Kunde]")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Land-Kunde als zweites Zeilenfeld setzen und Position festlegen
        Debug.Print "Land-Kunde als zweites Zeilenfeld setzen und Position festlegen"
        With wsDaten.PivotTables("pv_Daten").CubeFields("[KUNDE].[Land-Kunde]")
            .Orientation = xlRowField
            .Position = 2
        End With
        
        ' Produktgruppe Alternativ und Produkt Alternativ als Zeilenfelder hinzuf�gen
        Debug.Print "Produktgruppe Alternativ und Produkt Alternativ als Zeilenfelder hinzuf�gen"
        With wsDaten.PivotTables("pv_Daten").CubeFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ]")
            .Orientation = xlRowField
            .Position = 3
        End With
        
        With wsDaten.PivotTables("pv_Daten").CubeFields("[PRODUKT ALTERNATIV].[Produkt Alternativ]")
            .Position = 4
        End With
        
        ' * Produktgruppenebene alternativ 1 aufklappen
        Debug.Print "Produktgruppenebene alternativ 1 aufklappen"
        wsDaten.PivotTables("pv_Daten").PivotFields("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].[Produktgruppenebene alternativ 1]" _
            ).PivotItems("[PRODUKT ALTERNATIV].[Produktgruppe Alternativ].&[10000000100]"). _
            DrilledDown = True
        
        ' ZEIT-Monat als Zeilenfeld hinzuf�gen
        Debug.Print "ZEIT-Monat als Zeilenfeld hinzuf�gen"
        With wsDaten.PivotTables("pv_Daten").CubeFields("[ZEIT].[Monat]")
            .Position = 5
        End With
        
        ' Berechnete Ma�nahmen (Kosten) hinzuf�gen
        Debug.Print "Berechnete Ma�nahmen (Kosten) hinzuf�gen"
        measureHerstellkosten = "[Measures].[Herstellkosten] = [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB1 Rechnungsposition (bereinigt)]"
        measureLAP = "[Measures].[LAP] = [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB4 Rechnungsposition (bereinigt)]"
        measureWAP = "[Measures].[WAP] = [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB3 Rechnungsposition (bereinigt)]"

        ' Datenfelder hinzuf�gen
        Debug.Print "Datenfelder hinzuf�gen"
        wsDaten.PivotTables("pv_Daten").AddDataField wsDaten.PivotTables("pv_Daten").CubeFields("[Measures].[Rechnungswert (bereinigt)]")
        wsDaten.PivotTables("pv_Daten").AddDataField wsDaten.PivotTables("pv_Daten").CubeFields("[Measures].[Herstellkosten]")
        wsDaten.PivotTables("pv_Daten").AddDataField wsDaten.PivotTables("pv_Daten").CubeFields("[Measures].[LAP]")
        wsDaten.PivotTables("pv_Daten").AddDataField wsDaten.PivotTables("pv_Daten").CubeFields("[Measures].[WAP]")

        ' PivotTable-Layout anpassen
        Debug.Print "PivotTable-Layout anpassen"
        wsDaten.PivotTables("pv_Daten").RowAxisLayout xlTabularRow
        wsDaten.PivotTables("pv_Daten").RepeatAllLabels xlRepeatLabels
        
        ' Spaltenbreite anpassen
        Debug.Print "Spaltenbreite anpassen"
        Cells.Select
        Cells.EntireColumn.AutoFit
        
        ' Zelle A2 ausw�hlen
        Debug.Print "Zelle A2 ausw�hlen"
        Range("E2").Select
    End If
    Debug.Print "Ende von bt_Pvfuellen_Click"
End Sub

Private Sub bt_ImportSQL_Click()
    ImportExcelToSQLServer
End Sub

' Klicken auf Schaltfl�che "Berechnen"
Private Sub bt_Berechnen_VR_Click()
    'OptimizeVBA False
    Debug.Print "Starte DeleteOldSheetVR"
    DeleteOldSheetVR
    Debug.Print "Starte CreateVertriebsreport"
    CreateVertriebsreport
    Debug.Print "Starte TransferDataToVertriebsreport"
    TransferDataToVertriebsreport
    'UpdateButtonStatus
    'OptimizeVBA True
End Sub

Private Sub bt_loesch_wsDaten_click()


' L�schen des wsDaten-Blatts
Debug.Print "L�schen des wsDaten-Blatts"
wsDaten.Delete
' Wiederherstellen des wsDaten-Blatts
Debug.Print "Wiederherstellen des wsDaten-Blatts"
Set wsDaten = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
wsDaten.Name = "Daten"

Debug.Print "Starte UpdateButtonStatus"
UpdateButtonStatus
End Sub

Private Sub ComboBox_Monat_Change()
Debug.Print "Starte ComboBox_Monat_Change"
    If ThisWorkbook.Sheets("Settings").Range("A1").Value <> ComboBox_Monat.Value Then
        ' Aktuellen Filterwert aus der ComboBox speichern
        Debug.Print "Aktuellen Filterwert aus der ComboBox speichern"
        ThisWorkbook.Sheets("Settings").Range("A1").Value = ComboBox_Monat.Value
        Debug.Print "bt_Pvfuellen.Enabled = True"
        bt_Pvfuellen.Enabled = True
    Else
        Debug.Print "bt_Pvfuellen.Enabled = False"
        bt_Pvfuellen.Enabled = False
    End If

End Sub

Private Sub cboYear_Change()
Debug.Print "Start cboYear_Change"
    If Not initializing Then
        Debug.Print "cbJahr_Change wird aufgerufen"
        ' Aktualisieren des Schaltfl�chenstatus, wenn das Jahr ge�ndert wird
        Debug.Print "Aktualisieren des Schaltfl�chenstatus, wenn das Jahr ge�ndert wird"
        UpdateButtonStatus
        ' Aktualisieren des Monatsfilters, wenn das Jahr ge�ndert wird
        Debug.Print "Aktualisieren des Schaltfl�chenstatus, wenn das Jahr ge�ndert wird"
        'UpdateMonthFilter CInt(cboYear.Value), Month(DateValue("1 " & cboMonth.Value & " 2000"))
    End If
End Sub

Private Sub cboMonth_Change()
Debug.Print "cbMonat_Change wird aufgerufen"
    If Not initializing Then
        Debug.Print "cbMonat_Change wird aufgerufen"
        UpdateDatum1
        ' Aktualisieren des Schaltfl�chenstatus, wenn das Monat ge�ndert wird
        Debug.Print "Aktualisieren des Schaltfl�chenstatus, wenn der Monat ge�ndert wird"
        UpdateButtonStatus
        ' Aktualisieren des Monatsfilters, wenn das Monat ge�ndert wird
        Debug.Print "' Aktualisieren des Monatsfilters, wenn das Monat ge�ndert wird"
        'UpdateMonthFilter CInt(cboYear.Value), Month(DateValue("1 " & cboMonth.Value & " 2000"))
    End If
    Debug.Print "cbMonat_Change ende"
End Sub



' Diese Funktion nimmt einen Monatsnamen als Eingabe und gibt die entsprechende Monatsnummer zur�ck.
Function MonthNumber(MonthNameInput As String) As Integer
    Dim i As Integer
    Debug.Print "Function MonthNumber wird aufgerufen"
    
    ' Schleife von 1 bis 12 (Monatsnummern)
    For i = 1 To 12
        ' Vergleicht den Monatsnamen von der Funktion MonthName(i) mit dem eingegebenen Monatsnamen
        If MonthName(i) = MonthNameInput Then
            ' Wenn die Monatsnamen �bereinstimmen, gibt die Funktion die Monatsnummer zur�ck
            MonthNumber = i
            Exit Function
        End If
    Next i
End Function

Sub UpdateButtonStatus()
    ' Deklaration von Variablen f�r Jahr und Monat
    Dim iYear As Integer
    Dim iMonth As Integer
    
    Debug.Print "UpdateDatum1 wird aufgerufen"
    
    ' Zuweisen von Jahr und Monat aus den ComboBoxen
    iYear = Val(cboYear.Value)
    iMonth = MonthNumber(cboMonth.Value)
    Debug.Print "iYear = Val(cboYear.Value)" & iYear
    Debug.Print "iMonth = Val(cboYear.Value)" & iMonth
    
    ' �berpr�fen, ob Pivot-Tabelle existiert und Status der Schaltfl�chen entsprechend anpassen
    If Not PivotTableExists(wsDaten, "pv_Daten") Then
        bt_Verbinden.Enabled = True
        lblAktuelleAuswertung.Caption = "Gew�hlte Auswertung: " & cboMonth & cboYear & "."
    Else
        bt_Verbinden.Enabled = False

        If IsPivotTableFilled(wsDaten, iYear, iMonth) Then
            bt_Pvfuellen.Enabled = False
            bt_Berechnen_HK.Enabled = True
            Debug.Print "bt_Pvfuellen.Enabled= " & bt_Pvfuellen.Enabled
        Else
            bt_Pvfuellen.Enabled = True
            bt_Berechnen_HK.Enabled = False
        End If
    End If
    
    ' Daten-Worksheet aktivieren und Zelle A1 ausw�hlen
    With wsDaten
        .Activate
        .Range("E1").Select
    End With
    ' Aktualisieren des Labels zur Anzeige der ausgew�hlten Auswertung
    lblAktuelleAuswertung.Caption = "Aktuelle Auswertung: " & wsDaten.Cells(2, 6).Value
    Debug.Print "lblAktuelleAuswertung.Caption = Aktuelle Auswertung: " & wsDaten.Cells(2, 6).Value
End Sub
Sub UpdateDatum1()
    Dim Jahr As String
    Dim Monat As Integer
    Dim i As Integer

' Jahr aus der ComboBox cboYear zuweisen
Jahr = cboYear.Value
Debug.Print "Ausgew�hltes Jahr in UpdateDatum : " & Jahr
' �berpr�fen und zuweisen des Monats (als Zahl) aus der ComboBox cboMonth
For i = 1 To 12
    If MonthName(i) = cboMonth.Value Then
        Monat = i
        Exit For
    End If
Next i
' Formatieren des Datums und Speichern in der Variable datum1
Datum1 = Format(DateSerial(Jahr, Monat, 1), "yyyy-mm-ddTHH:MM:ss")
Debug.Print "Datum1  : " & Datum1
End Sub
Private Sub bt_END_Click()
    ThisWorkbook.Sheets("Settings").Range("A1").Value = cboMonth.Value
    ThisWorkbook.Sheets("Settings").Range("B1").Value = cboYear.Value
    Unload Me
End Sub

'Sub UpdateMonthFilter(yearVal As Integer, monthVal As Integer)
'    Dim pv As pivotTable
'    Dim pf As pivotField
'    Dim newFilter As String
'
'    If PivotTableExists(wsDaten, "pv_Daten") Then
'        Set pv = wsDaten.PivotTables("pv_Daten")
'        Set pf = pv.PivotFields("[ZEIT].[Monat].[Monat].[Quartal Des Jahres]")
'
'        ' Neuen Filter erstellen * Muss eine if abfrage rein die pr�ft ob das n�tig ist!
'        newFilter = Format(DateSerial(yearVal, monthVal, 1), "yyyy-mm")
'
'        ' Filter aktualisieren
'        On Error Resume Next
'        pf.ClearLabelFilters
'        pf.CurrentPage = newFilter
'        On Error GoTo 0
'    Else
'        MsgBox "Die Pivot-Tabelle 'pv_Daten' konnte nicht gefunden werden.", vbCritical, "Fehler"
'    End If
'End Sub
