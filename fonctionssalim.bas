Attribute VB_Name = "Fonctions"
Function derligne(colonne As Integer) As Integer

 derligne = Cells(Rows.Count, 1).End(xlUp).Row

End Function

Function dercolonne(ligne As Integer) As Integer

dercol = Cells(ligne, Columns.Count).End(xlToLeft).Column

End Function





