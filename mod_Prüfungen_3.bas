Attribute VB_Name = "mod_Prüfungen_3"
Public Function CheckAndCreateSheets() As Boolean
    Dim sheetName As Variant
    Dim sheetExists As Boolean
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim sheetNames As Variant
    
    Debug.Print "Funkton CheckAndCreateSheets aufgerufen"
    
    Set wb = ThisWorkbook

    ' Liste der benötigten Arbeitsblätter
    Debug.Print "Liste der benötigten Arbeitsblätter"
    sheetNames = Array("Daten", "HardKopy", "Vertriebsreport")

    ' Überprüfen, ob die Arbeitsblätter vorhanden sind, und bei Bedarf erstellen
    Debug.Print "Überprüfen, ob die Arbeitsblätter vorhanden sind, und bei Bedarf erstellen"
    For Each sheetName In sheetNames
        sheetExists = False
        For Each ws In wb.Worksheets
            If ws.Name = sheetName Then
                sheetExists = True
                Exit For
            End If
        Next ws

        If Not sheetExists Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = sheetName
        End If
    Next sheetName

    ' Funktion gibt True zurück, wenn alle Arbeitsblätter vorhanden oder erstellt wurden
    Debug.Print "Funktion gibt True zurück, wenn alle Arbeitsblätter vorhanden oder erstellt wurden"
    CheckAndCreateSheets = True
    Debug.Print "Ende der Funktion CheckAndCreateSheets"
End Function

Public Function CheckAndCreateListObjects() As Boolean
    Debug.Print "Start der Funktion CheckAndCreateListObjects"
    Dim lo As ListObject
    
    ' Prüfen und erstellen des ListObjects "tbl_VR"
    Debug.Print "Prüfen und erstellen des ListObjects tbl_VR"
    On Error Resume Next
    Set lo = wsVR.ListObjects("tbl_VR")
    On Error GoTo 0
    
    If lo Is Nothing Then
        Set lo = wsVR.ListObjects.Add(SourceType:=xlSrcRange, Source:=wsVR.Range("A1:B1"), XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2")
        lo.Name = "tbl_VR"
    End If

    ' Zuweisen des ListObjects "tbl_VR" zur globalen Variable
    Debug.Print "Zuweisen des ListObjects tbl_VR zur globalen Variable"
    Set tblVR = lo
    
    ' Prüfen und erstellen des ListObjects "tbl_HK"
    Debug.Print "Prüfen und erstellen des ListObjects tbl_HK"
    On Error Resume Next
    Set lo = wsHK.ListObjects("tbl_HK")
    On Error GoTo 0
    
    If lo Is Nothing Then
        Set lo = wsHK.ListObjects.Add(SourceType:=xlSrcRange, Source:=wsHK.Range("A1:B1"), XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2")
        lo.Name = "tbl_HK"
    End If
    
    ' Zuweisen des ListObjects "tbl_HK" zur globalen Variable
    Set tblHK = lo

    ' Funktion gibt True zurück, wenn alle ListObjects vorhanden oder erstellt wurden
    Debug.Print "Funktion gibt True zurück, wenn alle ListObjects vorhanden oder erstellt wurden - und Funktionsende"
    CheckAndCreateListObjects = True
End Function

' Die Funktion "TableExists" überprüft, ob eine Tabelle mit dem angegebenen Namen
' in einem bestimmten Arbeitsblatt vorhanden ist.
' Parameter:
' ws - Das Arbeitsblatt, in dem nach der Tabelle gesucht werden soll.
' tableName - Der Name der Tabelle, nach der gesucht wird.
' Rückgabewert:
' True, wenn die Tabelle im Arbeitsblatt gefunden wurde, andernfalls False.
Public Function TableExists(ws As Worksheet, tableName As String) As Boolean
    ' Debug-Ausgabe zur Verfolgung des Funktionsaufrufs
    Debug.Print "Start der Funktion TableExists"
    
    ' Variablendeklaration
    Dim tbl As ListObject

    ' Fehlerbehandlung: Fehler ignorieren und zum nächsten Codeabschnitt springen
    ' Dies ist nützlich, um zu verhindern, dass der Code bei einem Fehler abbricht
    ' und stattdessen einfach fortfährt.
    On Error Resume Next
    
    ' Versuche, die Tabelle mit dem angegebenen Namen im Arbeitsblatt zu finden
    ' und die Variable "tbl" darauf zu setzen.
    Set tbl = ws.ListObjects(tableName)
    
    ' Setze die Fehlerbehandlung zurück auf die Standardbehandlung (Fehler abfangen)
    On Error GoTo 0

    ' Wenn die Variable "tbl" gesetzt wurde (d. h., sie ist nicht Nothing),
    ' bedeutet dies, dass die Tabelle gefunden wurde, und die Funktion gibt True zurück.
    ' Andernfalls gibt die Funktion False zurück.
    TableExists = Not tbl Is Nothing
    Debug.Print "Ende der Funktion TableExists"
End Function


' Diese Funktion überprüft, ob die Arbeitsblätter "Daten" und "HardKopy" im Arbeitsbuch vorhanden sind.
Public Function CheckSheetsExist() As Boolean
    'Dim wsDaten As Worksheet
    'Dim wsHK As Worksheet
    Dim IsExpectedSheet As Boolean

    ' Fehlerbehandlung aktivieren, um Fehler beim Zuweisen der Arbeitsblätter zu vermeiden
    On Error Resume Next

    ' Fehlerbehandlung deaktivieren
    On Error GoTo 0

    ' Überprüfen, ob beide Arbeitsblätter zugewiesen wurden
    If Not wsDaten Is Nothing And Not wsHK Is Nothing Then
        ' Wenn beide Arbeitsblätter vorhanden sind, wird der Wert auf True gesetzt
        IsExpectedSheet = True
    Else
        ' Wenn eines oder beide Arbeitsblätter fehlen, wird der Wert auf False gesetzt
        IsExpectedSheet = False
    End If

    ' Die Funktion gibt den Wert von IsExpectedSheet zurück
    CheckSheetsExist = IsExpectedSheet
End Function

'Listet die erstellten Verbindungen zu den DatenFeldern im Cube
Public Sub ListPivotFieldNames()
    Dim pv As PivotTable
    Dim pf As PivotField
    
    ' Überprüfen, ob die Pivot-Tabelle "pv_Daten" vorhanden ist
    If PivotTableExists(wsDaten, "pv_Daten") Then
        Set pv = wsDaten.PivotTables("pv_Daten")
        
        For Each pf In pv.PivotFields
            Debug.Print pf.Name
        Next pf
    Else
        MsgBox "Die Pivot-Tabelle 'pv_Daten' konnte nicht gefunden werden.", vbCritical, "Fehler"
    End If
End Sub


' Diese Funktion prüft, ob die Pivot-Tabelle mit dem angegebenen Namen im Arbeitsblatt vorhanden ist
Public Function PivotTableExists(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    Dim pt As PivotTable
    Set pt = ws.PivotTables(tableName)
    ' Wenn kein Fehler auftritt, bedeutet dies, dass die Pivot-Tabelle gefunden wurde
    If Err.number = 0 Then
        PivotTableExists = True
    Else
        ' Andernfalls existiert die Pivot-Tabelle nicht
        PivotTableExists = False
    End If
    On Error GoTo 0
End Function

Public Function IsPivotTableFilled(ws As Worksheet, yearVal As Integer, monthVal As Integer) As Boolean
    Dim rng As Range
    Dim cellValue As String
    Dim expectedValue As String

    Debug.Print "Function IsPivotTableFilled wird Ausgeführt"

    ' Zugriff auf die Zelle F2 im Daten-Worksheet
    Set rng = ws.Cells(2, 6) ' 6 entspricht Spalte F

    ' Wert der Zelle extrahieren
    cellValue = rng.Value
    Debug.Print "zuweisung: cellValue = rng.Value = " & cellValue
    ' Erwarteten Wert basierend auf yearVal und monthVal generieren
    expectedValue = Format(DateSerial(yearVal, monthVal, 1), "mmmm") & " " & CStr(yearVal)
    Debug.Print "Erwarteten Wert basierend auf yearVal und monthVal generieren: expectedValue= " & expectedValue
    ' Debug-Ausgabe der Werte
    Debug.Print "Zelle F2 Wert: " & cellValue
    Debug.Print "Erwarteter Wert: " & expectedValue


    ' Überprüfen, ob der Wert der Zelle dem gewünschten Monat und Jahr entspricht
    If InStr(cellValue, expectedValue) > 0 Then
        IsPivotTableFilled = True
        Debug.Print "Überprüfen, ob der Wert der Zelle dem gewünschten Monat und Jahr entspricht: IsPivotTableFilled = " & IsPivotTableFilled
    Else
        IsPivotTableFilled = False
        Debug.Print "Überprüfen, ob der Wert der Zelle dem gewünschten Monat und Jahr entspricht: IsPivotTableFilled = " & IsPivotTableFilled
    End If
End Function


' Diese Funktion überprüft, ob ein Datenfeld mit dem angegebenen Feldnamen in der Pivot-Tabelle pt vorhanden ist.
Public Function DataFieldExists(pt As PivotTable, fieldName As String) As Boolean
    Dim pf As PivotField

    ' Aktiviert die Fehlerbehandlung, um Fehler beim Zuweisen des PivotFields zu vermeiden
    On Error Resume Next

    ' Versucht, das PivotField mit dem angegebenen Feldnamen zuzuweisen
    Set pf = pt.DataFields(fieldName)

    ' Wenn kein Fehler aufgetreten ist, bedeutet dies, dass das PivotField vorhanden ist
    If Err.number = 0 Then
        DataFieldExists = True
    Else
        ' Wenn ein Fehler aufgetreten ist, bedeutet dies, dass das PivotField nicht vorhanden ist
        DataFieldExists = False
    End If

    ' Deaktiviert die Fehlerbehandlung
    On Error GoTo 0
End Function


Function GetFaktor() As Double
    GetFaktor = ThisWorkbook.Worksheets("Settings").Range("C4").Value
End Function

Function ConvertDecimalSeparator(ByVal number As Double) As String
    Dim decimalSeparator As String
    decimalSeparator = Application.International(xlDecimalSeparator)
    ConvertDecimalSeparator = Replace(Format(number, "0.0000"), ".", decimalSeparator)
End Function

