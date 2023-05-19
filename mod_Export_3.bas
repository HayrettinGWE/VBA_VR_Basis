Attribute VB_Name = "mod_Export_3"
'Public Sub ImportExcelToSQLServer()
'    Dim cn As ADODB.connection
'    Dim rs As ADODB.Recordset
'    Dim ws As Worksheet
'    Dim r As Long
'
'    ' Verbindung zum OLAP herstellen
'    Set cn = New ADODB.connection
'    cn.Open "Provider=SQLOLEDB;Data Source=PEI2KGWEDB3;Initial Catalog=DW-GWE;Integrated Security=SSPI;"
'
'    ' Arbeitsblatt festlegen, das importiert werden soll
'    Set ws = ThisWorkbook.Sheets("Vertriebsreport")
'    r = 2 ' Startreihe (angenommen, dass Zeile 1 die Überschriften enthält)
'
'    ' Die Daten Zeile für Zeile importieren
'    Set rs = New ADODB.Recordset
'    rs.Open "Vertriebsdaten", cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    Do While Not isEmpty(ws.Cells(r, 1))
'        rs.AddNew
'        rs.Fields("Kunden_Nr") = ws.Cells(r, 1)
'        rs.Fields("Kunde") = ws.Cells(r, 2)
'        rs.Fields("Leander_Code") = ws.Cells(r, 3)
'        rs.Fields("PG_Ebene") = ws.Cells(r, 4)
'        rs.Fields("PG") = ws.Cells(r, 5)
'        rs.Fields("PGA_Nr") = ws.Cells(r, 6)
'        rs.Fields("PGA") = ws.Cells(r, 7)
'        rs.Fields("Monat") = ws.Cells(r, 8)
'        rs.Fields("Umsatz") = ws.Cells(r, 9)
'        rs.Fields("HK") = ws.Cells(r, 10)
'        rs.Fields("LAP_Lager") = ws.Cells(r, 11)
'        rs.Fields("WAP_Werk") = ws.Cells(r, 12)
'        rs.Fields("Kosten_DB1_Transport") = ws.Cells(r, 13)
'        rs.Fields("Marge_DB1") = ws.Cells(r, 14)
'        rs.Fields("Marge_DB1_Prozent") = ws.Cells(r, 15)
'        rs.Fields("Zuschlaege_DB3") = ws.Cells(r, 16)
'        rs.Fields("Kosten_DB3") = ws.Cells(r, 17)
'        rs.Fields("Marge_DB3") = ws.Cells(r, 18)
'        rs.Fields("Marge_DB3_Prozent") = ws.Cells(r, 19)
'        rs.Fields("AD_MA") = ws.Cells(r, 20)
'        rs.Fields("PE_Haendler") = ws.Cells(r, 21)
'        rs.Fields("EinbauWerkZ") = ws.Cells(r, 22)  ' Noch hinzufügen und nochmal hochladen!!!
'        rs.Fields("Gebiet") = ws.Cells(r, 23)
'        rs.Fields("IC_Gesellschaft") = ws.Cells(r, 24) ' Hier habe ich die Spalte korrigiert
'        rs.Update
'        r = r + 1
'    Loop
'
'    ' Verbindung schließen
'    rs.Close
'    cn.Close
'    Set rs = Nothing
'End Sub

Public Sub ImportExcelToSQLServer()
    Dim cn As ADODB.connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim r As Long

    ' Verbindung zum OLAP herstellen
    Set cn = New ADODB.connection
    cn.Open "Provider=SQLOLEDB;Data Source=PEI2KGWEDB3;Initial Catalog=DW-GWE;Integrated Security=SSPI;"

    ' Arbeitsblatt festlegen, das importiert werden soll
    Set ws = ThisWorkbook.Sheets("ExpDB")
    r = 2 ' Startreihe (angenommen, dass Zeile 1 die Überschriften enthält)

    ' Die Daten Zeile für Zeile importieren
    Set rs = New ADODB.Recordset
    rs.Open "Vertriebsdaten", cn, adOpenStatic, adLockOptimistic, adCmdTable

    Do While Not IsEmpty(ws.Cells(r, 1))
        rs.AddNew
        rs.Fields("DATUM") = ws.Cells(r, 1)
        rs.Fields("KDNR") = ws.Cells(r, 2)
        rs.Fields("AGGREG_NR") = ws.Cells(r, 3)
        rs.Fields("ANR") = ws.Cells(r, 4)
        rs.Fields("RG_WERT_BEREINIGT") = ws.Cells(r, 5)
        rs.Fields("HK") = ws.Cells(r, 6)
        rs.Fields("LAP") = ws.Cells(r, 7)
        rs.Fields("WAP") = ws.Cells(r, 8)
        rs.Fields("Kosten_DB1_Transport") = ws.Cells(r, 9)
        rs.Fields("Marge_DB1") = ws.Cells(r, 10)
        rs.Fields("Marge_DB1_Prozent") = ws.Cells(r, 11)
        rs.Fields("Zuschlaege_DB3") = ws.Cells(r, 12)
        rs.Fields("Kosten_DB3") = ws.Cells(r, 13)
        rs.Fields("Marge_DB3") = ws.Cells(r, 14)
        rs.Fields("Marge_DB3_Prozent") = ws.Cells(r, 15)
        rs.Fields("AD_MA") = ws.Cells(r, 16)
        rs.Fields("Gebiet") = ws.Cells(r, 17)
        rs.Fields("PE_Haendler") = IIf(ws.Cells(r, 18) = "Ja", 1, 0)
        rs.Fields("EinbauWerkZ") = IIf(ws.Cells(r, 19) = "Ja", 1, 0)
        rs.Fields("IC_Gesellschaft") = IIf(ws.Cells(r, 20) = "Ja", 1, 0)
        rs.Update
        r = r + 1
    Loop

    ' Verbindung schließen
    rs.Close
    cn.Close
    Set rs = Nothing
End Sub

