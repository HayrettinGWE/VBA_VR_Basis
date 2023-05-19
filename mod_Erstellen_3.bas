Attribute VB_Name = "mod_Erstellen_3"
' Diese Funktion leert den Inhalt der angegebenen Tabelle im angegebenen Arbeitsblatt ab Zeile 2.
Public Function ClearTableContents(sheetName As String, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Durchlaufen aller Arbeitsblätter
    For Each ws In Sheets
        ' Prüfen, ob das Arbeitsblatt dem übergebenen Namen entspricht
        If ws.Name = sheetName Then
            ' Durchlaufen aller Tabellen im Arbeitsblatt
            For Each tbl In ws.ListObjects
                ' Prüfen, ob die Tabelle dem übergebenen Namen entspricht
                If tbl.Name = tableName Then
                    ' Alle Einträge in der Tabelle ab Zeile 2 löschen
                    tbl.DataBodyRange.Rows(2).Resize(tbl.ListRows.Count - 1).ClearContents
                    Exit For ' Beendet die Schleife, wenn die Tabelle gefunden wurde
                End If
            Next tbl
            Exit For ' Beendet die Schleife, wenn das Arbeitsblatt gefunden wurde
        End If
    Next ws
End Function

