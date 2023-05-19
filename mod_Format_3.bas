Attribute VB_Name = "mod_Format_3"

' Diese Prozedur ändert das Zahlenformat einer Spalte in einem Arbeitsblatt in Euroformat.
Sub SinEurFormat(ByVal ws As Worksheet, ByVal columnLetter As String)
    Dim col As Range

    ' Die entsprechende Spalte im Arbeitsblatt auswählen
    Set col = ws.Columns(columnLetter)
    ' Das Zahlenformat der Spalte auf Euro setzen
    col.NumberFormat = "#,##0.00 €"
End Sub

' Diese Prozedur ändert das Zahlenformat einer Spalte in einem Arbeitsblatt in Prozentformat.
Public Sub SetPercentFormat(ByVal ws As Worksheet, ByVal columnLetter As String)
    Dim col As Range

    ' Die entsprechende Spalte im Arbeitsblatt auswählen
    Set col = ws.Columns(columnLetter)
    ' Das Zahlenformat der Spalte auf Prozent setzen
    col.NumberFormat = "0.00%"
End Sub

' Diese Prozedur ändert das Zahlenformat einer Spalte in einem Arbeitsblatt in Textformat.
Public Sub SetTextFormat(ByVal ws As Worksheet, ByVal columnLetter As String)
    Dim col As Range

    ' Die entsprechende Spalte im Arbeitsblatt auswählen
    Set col = ws.Columns(columnLetter)
    ' Das Zahlenformat der Spalte auf Text setzen
    col.NumberFormat = "@"
End Sub

' Dieser Sub löscht ein Arbeitsblatt namens "Vertriebsreport", falls es vorhanden ist.
Public Sub DeleteOldSheetVR()
    Dim blatt As Worksheet
    ' Durchlaufen aller Arbeitsblätter
    For Each blatt In Sheets
        ' Prüfen, ob das Arbeitsblatt "HardKopy" heißt
        If blatt.Name = "Vertriebsreport" Then
            ' "HardKopy" auswählen
            Sheets("Vertriebsreport").Select
            ' Das ausgewählte Arbeitsblatt löschen
            ActiveWindow.SelectedSheets.Delete
        End If
    Next blatt
End Sub

' Prozedur zum Formatieren der Pivot-Tabelle
Public Sub FormatPivotTable()
    ' Pivot-Tabelle formatieren
    With ActiveSheet.PivotTables("pv_Daten")
        ' Spaltensumme aktivieren
        
        .ColumnGrand = True
        ' Automatische Formatierung aktivieren
        .HasAutoFormat = True
        ' Fehlermeldungen nicht anzeigen
        .DisplayErrorString = False
        ' Leere Zellen anzeigen
        .DisplayNullString = True
        ' Drilldown aktivieren
        .EnableDrilldown = True
        ' Fehlermeldung festlegen
        .ErrorString = ""
        ' Beschriftungen zusammenführen deaktivieren
        .MergeLabels = False
        ' Text für leere Zellen festlegen
        .NullString = ""
        ' Reihenfolge der Seitenfelder festlegen
        .PageFieldOrder = 2
        ' Anzahl der Spalten im Seitenfeld festlegen
        .PageFieldWrapCount = 0
        ' Formatierung beibehalten aktivieren
        .PreserveFormatting = True
        ' Zeilensumme aktivieren
        .RowGrand = True
        ' Titel drucken deaktivieren
        .PrintTitles = False
        ' Elemente auf jeder gedruckten Seite wiederholen
        .RepeatItemsOnEachPrintedPage = True
        ' Gesamtannotation aktivieren
        .TotalsAnnotation = True
        ' Einzug für kompakte Zeilen festlegen
        .CompactRowIndent = 1
        ' Visuelle Summen deaktivieren
        .VisualTotals = False
        ' Raster-Drop-Zonen deaktivieren
        .InGridDropZones = False
        ' Feldbeschriftungen anzeigen
        .DisplayFieldCaptions = True
        ' Member-Eigenschafts-Tooltips anzeigen
        .DisplayMemberPropertyTooltips = True
        ' Kontext-Tooltips anzeigen
        .DisplayContextTooltips = True
        ' Bohrungsindikatoren anzeigen
        .ShowDrillIndicators = True
        ' Bohrungsindikatoren nicht drucken
        .PrintDrillIndicators = False
        ' Leere Zeilen nicht anzeigen
        .DisplayEmptyRow = False
        ' Leere Spalten nicht anzeigen
        .DisplayEmptyColumn = False
        ' Mehrfachfilter nicht zulassen
        .AllowMultipleFilters = False
        ' Benutzerdefinierte Listen für die Sortierung verwenden
        .SortUsingCustomLists = True
        ' Sofortige Elemente anzeigen
        .DisplayImmediateItems = True
        ' Berechnete Mitglieder anzeigen
        .ViewCalculatedMembers = True
        ' Schreibgeschützt deaktivieren
        .EnableWriteback = False
        ' Wertezeile nicht anzeigen
        .ShowValuesRow = False
        ' Berechnete Mitglieder in Filtern anzeigen
        .CalculatedMembersInFilters = True
        ' Zeilenachsenlayout auf kompakt setzen
        .RowAxisLayout xlCompactRow
    End With
End Sub
