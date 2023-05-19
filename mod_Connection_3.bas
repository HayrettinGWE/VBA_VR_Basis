Attribute VB_Name = "mod_Connection_3"

' Funktion zum Herstellen einer Verbindung zum OLAP
Public Function ConnectToOLAP() As ADODB.connection
    Dim cn As ADODB.connection
    Set cn = New ADODB.connection

    ' Verbindung mit dem OLAP-Server herstellen
    cn.Open "Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;Data Source=PEI2KGWEDB3;Update Isolation Level=2;Initial Catalog=DW-D21"
    Set ConnectToOLAP = cn
End Function


' Hauptprozedur zum Verbinden und Erstellen der Pivot-Tabelle
Public Sub ConnectAndCreatePivotTable()
    Dim cn As ADODB.connection
    ' Verbindung zum OLAP herstellen
    Set cn = ConnectToOLAP()

    ' �berpr�fen, ob die ben�tigten Arbeitsbl�tter vorhanden sind
    If CheckSheetsExist() Then
        ' Pr�fung, ob Verbindungen existieren
        If Not ActiveWorkbook.Connections.Count = 0 Then
            For Each connection In ActiveWorkbook.Connections
                If connection.Name = "PEI2KGWEDB3 DW-D21 Vertrieb" Then
                    With connection.OLEDBConnection
                        .connection = Array("OLEDB;Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;Data Source=PEI2KGWEDB3;Update Isolation Level=2;Initial C", "atalog=DW-D21")
                        .MaxDrillthroughRecords = 1048576
                    End With
                    Exit For
                End If
            Next connection
        End If

        ' Verbindung zum Data Warehouse hinzuf�gen
        Workbooks("Schnittstelle Datawearhaous_v4.xlsm").Connections.Add2 "PEI2KGWEDB3 DW-D21 Vertrieb", "", Array("OLEDB;Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;Data Source=PEI2KGWEDB3;Update Isolation Level=2;Initial C", "atalog=DW-D21"), "Vertrieb", 1

        On Error Resume Next

        ' Pr�fen, ob bereits eine Pivot-Tabelle vorhanden ist
        If ActiveSheet.PivotTables.Count = 0 Then
            ' Pivot-Tabelle erstellen
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:=ActiveWorkbook.Connections("PEI2KGWEDB3 DW-D21 Vertrieb"), Version:=7).CreatePivotTable TableDestination:="Daten!R1C1", tableName:="pv_Daten", DefaultVersion:=7

            ' �berpr�fen, ob die Pivot-Tabelle erfolgreich erstellt wurde
            If ActiveSheet.PivotTables("pv_Daten") Is Nothing Then
                MsgBox "Das PivotTable konnte nicht erstellt werden."
            Else
                ' Pivot-Tabelle formatieren
                Call FormatPivotTable
            End If
        Else
            MsgBox "Das PivotTable ist bereits vorhanden."
        End If
        On Error GoTo 0
        
          ' Berechnete Ma�nahmen hinzuf�gen
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = cn
        
        ' Berechnete Ma�nahmen als MDX-Abfragen hinzuf�gen
        measureHerstellkosten = "CALCULATE; CREATE MEMBER CURRENTCUBE.[Measures].[Herstellkosten] AS [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB1 Rechnungsposition (bereinigt)]"
        measureLAP = "CALCULATE; CREATE MEMBER CURRENTCUBE.[Measures].[LAP] AS [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB4 Rechnungsposition (bereinigt)]"
        measureWAP = "CALCULATE; CREATE MEMBER CURRENTCUBE.[Measures].[WAP] AS [Measures].[Rechnungswert (bereinigt)] - [Measures].[DB3 Rechnungsposition (bereinigt)]"
        
        ' Berechnete Ma�nahmen in den OLAP-W�rfel einf�gen
'        cmd.CommandText = measureHerstellkosten
'        cmd.Execute
'
'        cmd.CommandText = measureLAP
'        cmd.Execute
'
'        cmd.CommandText = measureWAP
'        cmd.Execute
        
        ' OLAP-Verbindung schlie�en
        cn.Close

    Else
        MsgBox "Bitte stellen Sie sicher, dass die Arbeitsbl�tter 'Daten' und 'HardKopy' vorhanden sind."
    End If
End Sub

