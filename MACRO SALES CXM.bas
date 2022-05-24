Attribute VB_Name = "Module81"



Private Sub CommandButton1_Click()
Call splitcommande
End Sub
Sub import_PSO()

Dim derligne As Integer
Dim pso, wb As Workbook
Dim ws As Worksheet
Dim nom As Variant
Dim feuille As String

Set wb = Workbooks("FICHIER SALES  CXM.xlsm")
Set ws = wb.Worksheets(2)
wb.Activate
derlignelwb = lastrow(3)
nom = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
If nom <> False Then
    Set pso = Application.Workbooks.Open(nom)
Else
    Exit Sub
End If

pso.Activate
feuille = InputBox("Quelle feuille voulez-vous importer ?")

Sheets(feuille).Activate
derligne = lastrow(1)
dercol = lastCol(1)

'CC des numéros SAP'
pso.Activate
Sheets(feuille).Activate
Range(Cells(3, 1), Cells(derligne, 1)).Copy
wb.Activate
'ws.Cells(derlignelwb + 1, 4).Select
Cells(derlignelwb + 1, 4).PasteSpecial Paste:=xlPasteValues

'client name
pso.Activate
Sheets(feuille).Activate
Range(Cells(3, 3), Cells(derligne, 3)).Copy
wb.Activate
Cells(derlignelwb + 1, 6).PasteSpecial Paste:=xlPasteValues

'montant
pso.Activate
Sheets(feuille).Activate
Range(Cells(3, 4), Cells(derligne, 4)).Copy
wb.Activate
Cells(derlignelwb + 1, 22).PasteSpecial Paste:=xlPasteValues


'vendeurs
pso.Activate
Sheets(feuille).Activate
Range(Cells(3, 9), Cells(derligne, 9)).Copy
wb.Activate
Cells(derlignelwb + 1, 3).PasteSpecial Paste:=xlPasteValues

'date
wb.Activate
datepso = InputBox("Saisir une date pour les PSO")
For i = lastrow(5) + 1 To lastrow(3)
    Cells(i, 5) = datepso
Next

'pays
Dim code As String
pays = InputBox("Taper FR ou BNL")
 Select Case pays
    Case Is = "FR"
    code = "QUADFrance(FR00)"
    Case Is = "BNL"
    code = "QUADBenelux(CH06)"
End Select
For i = lastrow(2) + 1 To lastrow(3)
    Cells(i, 2) = code
Next


'uniformisation nom vendeur
wb.Activate
For i = 10 To lastrow(3)
    If Cells(i, 3) = "Sébastien" Then
        Cells(i, 3) = "Sebastien"
    End If
    Cells(i, 3).Value = Replace(Cells(i, 3), " ", "")
Next
MsgBox "Import PSO terminer, veuillez supprimer les éventuelles lignes vides"

End Sub

Sub load_data()
Dim derligne As Integer
Dim sfdc, wb As Workbook
Dim ws As Worksheet
Dim nom As Variant


Set wb = Workbooks("FICHIER SALES  CXM.xlsm")
Set ws = wb.Worksheets(2)
wb.Activate
derlignelwb = lastrow(2)
nom = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
If nom <> False Then
    Set sfdc = Application.Workbooks.Open(nom)
Else
    Exit Sub
End If

sfdc.Activate
Sheets(1).Activate
derligne = lastrow(1)
dercol = lastCol(1)

'remplacement des points par des virgules'
For l = 2 To derligne
    For c = 1 To dercol
                Cells(l, c).Value = Replace(Cells(l, c), ".", ",")
                Cells(l, c).Value = Replace(Cells(l, c), " ", "")
                'Cells(l, c).Font.ColorIndex = 14
            
    Next
Next

'CC CV de la colonne sale organization'
Range(Cells(2, 2), Cells(derligne, 2)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC CV VENDEUR'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 28), Cells(derligne, 28)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 3).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC des numéros SAP'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 13), Cells(derligne, 13)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 4).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC des  DATES'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 24), Cells(derligne, 24)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 5).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC des  Noms client'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 4), Cells(derligne, 4)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 6).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC des Types d'affaires'
'sfdc.Activate
'Sheets(1).Activate
'Range(Cells(2, 28), Cells(derligne, 28)).Copy
'Workbooks("FICHIER SALES  CXM.xlsx").Activate
'ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC  Licences '
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 6), Cells(derligne, 6)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 7).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC  Maintenance '
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 8), Cells(derligne, 8)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 8).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC  Souscription '
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 10), Cells(derligne, 10)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 9).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC  Souscription '


'Mise en forme'
wb.Activate
For i = 10 To lastrow(2)
    For y = 7 To 9
        Cells(i, y) = CCur(Cells(i, y))
    Next
Next
     
'CCur()'
End Sub

Sub calculs()
Dim obj, psobj, boost, boost1, boost2 As Currency
Dim dico
Dim mois As String
Dim taux59, taux79, taux100, taux100max, saas, psorate As Currency

'définition des objectifs
obj = Cells(3, 5)
psobj = Cells(4, 8)
boost1 = 1.143
boost2 = 2.143

'définition des taux de commission
psorate = Cells(4, 9)
saas = Cells(4, 12)
taux59 = Cells(4, 6)
taux79 = Cells(5, 6)
taux100 = Cells(6, 6)
taux100max = Cells(7, 6)


'Worksheets("Monthly Commissions").Activate'

'Calcul de la colonne Total revenue'
For i = 10 To lastrow(2)
    Cells(i, 14) = Cells(i, 7) + Cells(i, 8) + Cells(i, 13)
Next

'Calculs PSO'
For i = 10 To lastrow(2)
    'Calcul de la colonne cumulative revenue  et cumulative pso'
    Cells(i, 15) = WorksheetFunction.Sum(Range(Cells(10, 14), Cells(i, 14)))
    Cells(i, 15).NumberFormat = "0.00€"
    
    Cells(i, 23) = WorksheetFunction.Sum(Range(Cells(10, 22), Cells(i, 22)))
    'R/O PSO
    Cells(i, 24) = (Cells(i, 23) / psobj)
    Cells(i, 24).STYLE = "Percent"
    Cells(i, 24).NumberFormat = "0.00%"
    
    'Commission rate pso
    Cells(i, 25) = psorate / 1
    Cells(i, 25).STYLE = "Percent"
    Cells(i, 25).NumberFormat = "0.00%"
    
    'commission pso
    Cells(i, 26) = psorate * Cells(i, 22)
    Cells(i, 26).NumberFormat = "0.00€"
Next


'Calcul des souscriptions boostées'
For i = 10 To lastrow(2)
    
    If Cells(i, 10) > 1 Then
    
        Cells(i, 13) = Cells(i, 9) * boost2
    Else
        Cells(i, 13) = Cells(i, 9) * boost1
        
    End If
    Cells(i, 13) = CCur(Cells(i, 13))
Next


'Calcul des R/O'
For i = 10 To lastrow(2)
    Cells(i, 16) = (Cells(i, 15) / obj)
    Cells(i, 16).STYLE = "Percent"
    Cells(i, 16).NumberFormat = "0.00%"
Next


'Création d'un dictionnaire des Mois commerciaux
Set dico = CreateObject("Scripting.Dictionary")
With dico
    .Add "01", "M12"
    .Add "02", "M1"
    .Add "03", "M2"
    .Add "04", "M3"
    .Add "05", "M4"
    .Add "06", "M5"
    .Add "07", "M6"
    .Add "08", "M7"
    .Add "09", "M8"
    .Add "10", "M9"
    .Add "11", "M10"
    .Add "12", "M11"
End With

'application du mois commercial à chaque case
For i = 10 To lastrow(2)
mois = Right(Left(Cells(i, 5), 5), 2)
Cells(i, 1) = dico.Item(mois)
Next


'Calculs des taux de commission de base

For i = 10 To lastrow(2)
    Select Case Cells(i, 16)
        Case Is < 59 / 100
            Cells(i, 17) = taux59
        Case 59 / 100 To 79 / 100
            Cells(i, 17) = taux79
        Case 79 / 100 To 100 / 100
            Cells(i, 17) = taux100
        Case Is < 100 / 100
            Cells(i, 17) = taux100max
    End Select
Cells(i, 17).NumberFormat = "0.00%"

'calcul du commissionnemet en €
Cells(i, 18) = Cells(i, 17) * Cells(i, 14)
Cells(i, 18).NumberFormat = "0.00€"
Next

'Calcul Saas kicker commission
For i = 10 To lastrow(2)
    Cells(i, 21) = Cells(i, 12) * saas
    Cells(i, 21).NumberFormat = "0.00€"
Next

'total commissions
For i = 10 To lastrow(2)
    Cells(i, 27) = Cells(i, 18) + Cells(i, 20) + Cells(i, 21) + Cells(i, 26)
    Cells(i, 27).NumberFormat = "0.00€"
Next

End Sub

Sub splitcommande()

Dim palier, ponderation As Double
Dim obj, target, CA As Currency
Dim licence, maint, abon As Currency
Dim ligne As Integer

obj = Cells(3, 5)
palier = Application.InputBox("Quel palier a été atteint? Entrer 60, 80 ou 100")
If palier = False Then
    Exit Sub
End If

palier = palier / 100
target = obj * palier

ligne = Application.InputBox("Quel est le numéro de ligne de la commande qui dépasse un palier?")
If ligne = False Then
    Exit Sub
End If


CA = Cells(ligne, 14)

diff = target - Cells(ligne - 1, 15)
ponderation = diff / CA
licence = Cells(ligne, 7)
maint = Cells(ligne, 8)
abon = Cells(ligne, 9)

'Insertion d'une ligne identique à celle à split'
Rows(ligne + 1).Insert
Rows(ligne).Copy
Rows(ligne + 1).Select
Selection.PasteSpecial Paste:=xlPasteValues
Rows(ligne).Font.ColorIndex = 14

'modification des lignes de CA avant palier'
licence = ponderation * licence
Cells(ligne, 7) = licence
Cells(ligne, 7).Font.ColorIndex = 13

maint = ponderation * maint
Cells(ligne, 8) = maint
Cells(ligne, 8).Font.ColorIndex = 13

'modification du montant de abo si non nul'
If Cells(ligne, 10) <> 0 Then

    If Cells(ligne, 10) > 1 Then
    
            Cells(ligne, 9) = Cells(ligne, 13) / Cells(5, 13)
        Else
            Cells(ligne, 9) = Cells(ligne, 13) / Cells(4, 13)
    
    End If
Cells(ligne, 9).Font.ColorIndex = 13
End If

'modification des lignes de CA après palier'
Cells(ligne, 7).Offset(1, 0).Select
Selection = licence * (1 - ponderation)
Selection.Font.ColorIndex = 13

Cells(ligne, 8).Offset(1, 0).Select
Selection = maint * (1 - ponderation)
Selection.Font.ColorIndex = 13

Cells(ligne, 9).Offset(1, 0).Select
Selection = Cells(ligne, 9) * (1 - ponderation)
Selection.Font.ColorIndex = 13

Call calculs

End Sub

Sub dispatch()

Dim start As Currency
Dim name As String
Dim wb As Workbook
Dim liste_name As Object
Dim vendeurs As Variant
Dim derligne As Integer

derligne = Cells(Rows.Count, 2).End(xlUp).Row

Set liste_name = CreateObject("System.Collections.ArrayList")

start = Application.InputBox("Ligne de démararrage ?")

Worksheets(2).Activate
For i = start To lastrow(2)
    name = Cells(i, 3)
    Range(Cells(i, 1), Cells(i, 23)).Copy
    
    Worksheets(name).Cells(lastrow(2, name) + 1, 1).PasteSpecial Paste:=xlPasteValues
    
    'Cells(11, 1).Select
    'Range("A10").Select
    
    'Selection.PasteSpecial Paste:=xlPasteValues
    If liste_name.contains(name) Then
    Else
    liste_name.Add name
    End If

        
Next
'recalcul pour chaque onglets
For Each vendeurs In liste_name
    Worksheets(vendeurs).Activate
    Call calculs
Next

End Sub

