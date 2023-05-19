VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim answer As Variant

    ' Eine MsgBox anzeigen, um den Benutzer zu fragen, ob das Vertriebsreport-Programm gestartet werden soll
    answer = MsgBox("Programm Vertriebsreport starten?", vbYesNo + vbQuestion, "Vertriebsreport-Programm")

    ' Prüfen, ob der Benutzer auf "Ja" geklickt hat
    If answer = vbYes Then
        ' Die UserForm anzeigen
        uf_Basis_3.Show
    End If
End Sub
