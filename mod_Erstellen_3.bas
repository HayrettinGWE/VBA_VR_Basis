Attribute VB_Name = "mod_Erstellen_3"
' Diese Funktion leert den Inhalt der angegebenen Tabelle im angegebenen Arbeitsblatt ab Zeile 2.
Public Function ClearTableContents(sheetName As String, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Durchlaufen aller Arbeitsbl�tter
    For Each ws In Sheets
        ' Pr�fen, ob das Arbeitsblatt dem �bergebenen Namen entspricht
        If ws.Name = sheetName Then
            ' Durchlaufen aller Tabellen im Arbeitsblatt
            For Each tbl In ws.ListObjects
                ' Pr�fen, ob die Tabelle dem �bergebenen Namen entspricht
                If tbl.Name = tableName Then
                    ' Alle Eintr�ge in der Tabelle ab Zeile 2 l�schen
                    tbl.DataBodyRange.Rows(2).Resize(tbl.ListRows.Count - 1).ClearContents
                    Exit For ' Beendet die Schleife, wenn die Tabelle gefunden wurde
                End If
            Next tbl
            Exit For ' Beendet die Schleife, wenn das Arbeitsblatt gefunden wurde
        End If
    Next ws
End Function

