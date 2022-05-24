Attribute VB_Name = "Module8"



Private Sub CommandButton1_Click()
Call splitcommande
End Sub

Sub load_data_EMEA()
Dim derligne As Integer
Dim sfdc, wb As Workbook
Dim ws As Worksheet
Dim nom As Variant


Set wb = ThisWorkbook
Set ws = wb.Worksheets(1)
wb.Activate
derlignelwb = lastrow(1)
MsgBox "Veuillez ouvrir le rapport SFDC EMEA BOOKING"
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
                Cells(l, c).Font.ColorIndex = 14
            
    Next
Next

'CC CV de la colonne sale organization, code pays
Range(Cells(2, 29), Cells(derligne, 29)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC CV CURRENCY
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 4), Cells(derligne, 4)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 3).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC des numéros SAP'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 12), Cells(derligne, 12)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 4).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC des  DATES'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 23), Cells(derligne, 23)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 5).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC des  Noms client'
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 3), Cells(derligne, 3)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 6).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC  Licences '
'    sfdc.Activate
'    Sheets(1).Activate
'    Range(Cells(2, 5), Cells(derligne, 5)).Copy
'    wb.Activate
'    ws.Cells(derlignelwb + 1, 7).Select
'    ActiveCell.PasteSpecial Paste:=xlPasteValues
 '
'CC  Maintenance '
'sfdc.Activate
'Sheets(1).Activate
'Range(Cells(2, 7), Cells(derligne, 7)).Copy
'wb.Activate
'ws.Cells(derlignelwb + 1, 8).Select
'ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC Annual Souscription '
'sfdc.Activate
'Sheets(1).Activate
'Range(Cells(2, 9), Cells(derligne, 9)).Copy
'wb.Activate
'ws.Cells(derlignelwb + 1, 9).Select
'ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC DURATION
'sfdc.Activate
'Sheets(1).Activate
'Range(Cells(2, 15), Cells(derligne, 15)).Copy
'wb.Activate
'ws.Cells(derlignelwb + 1, 10).Select
'ActiveCell.PasteSpecial Paste:=xlPasteValues


'CC de la colonne converted amount
sfdc.Activate
Sheets(1).Activate
Range(Cells(2, 26), Cells(derligne, 26)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 14).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues


'Mise en forme'
wb.Activate
For i = 10 To lastrow(2)
    For y = 7 To 9
        Cells(i, y) = CCur(Cells(i, y))
    Next
Next
     
    
moisco = Application.InputBox("De quel mois commercial s'agit-il? (M1, M2...)")
     
wb.Activate
For i = derlignelwb + 1 To lastrow(6)
    'Cells(i, 5).Format = "dd/mm/aaaa"
    Cells(i, 1) = moisco
    For y = 7 To 9
      Cells(i, y) = CCur(Cells(i, y))
    Next

Next
End Sub

Sub calculs_EMEA()
Dim boost, boost1, boost2 As Currency
Dim obj, objloc As Double
Dim dico
Dim mois As String
Dim locale As Boolean
Dim taux100, taux100max, tx100local, tx100localmax, psorate As Currency


'définition des objectifs
obj = Cells(3, 5)
objloc = Cells(5, 4)
boost1 = 1.143
boost2 = 2.143

'définition des taux de commission

saas = Cells(4, 12)

tx100localmax = Cells(4, 6)
tx100local = Cells(5, 6)

taux100max = Cells(6, 6)
taux100 = Cells(7, 6)


'calcul des subscriptions
For i = 10 To lastrow(2)
    If Cells(i, 9) <> 0 And Cells(i, 8) = 0 Then
    Cells(i, 7) = 0
    End If
Next

'Calcul de la colonne Total revenue'
For i = 10 To lastrow(2)

   ' Cells(i, 14) = Cells(i, 7) + Cells(i, 8) + Cells(i, 13)
    Cells(i, 14) = CCur(Cells(i, 14))
    Cells(i, 14).NumberFormat = "0.00€"
Next


'Calcul des souscriptions boostées'
For i = 10 To lastrow(2)
    
    If Cells(i, 10) > 1 Then
    
        Cells(i, 13) = Cells(i, 9) * 1
    Else
        Cells(i, 13) = Cells(i, 9) * 1
        
    End If
    Cells(i, 13) = CCur(Cells(i, 13))
Next


'calcul du cumul
For i = 10 To lastrow(2)
    'Calcul de la colonne cumulative revenue  et cumulative pso'
    Cells(i, 15) = WorksheetFunction.Sum(Range(Cells(10, 14), Cells(i, 14)))
    Cells(i, 17) = Application.Sum(Application.SumIfs(Range(Cells(10, 14), Cells(i, 14)), Range(Cells(10, 2), Cells(i, 2)), Array("QUADFrance(FR00)", "QUADBenelux(CH06)", "QUADDenmark(DK00)")))
    Cells(i, 17).NumberFormat = "0.00€"
    'Array("QUADFrance(FR00)", "QUADBenelux(CH06)"))
Next



'Calcul des R/O'
For i = 10 To lastrow(2)
    Cells(i, 16) = (Cells(i, 15) / obj)
    Cells(i, 16).Style = "Percent"
    Cells(i, 16).NumberFormat = "0.00%"
    
    If objloc = 0 Then
        Cells(i, 18) = 0
        
    Else
        Cells(i, 18) = (Cells(i, 17) / objloc)
    End If
    
    Cells(i, 18).Style = "Percent"
    Cells(i, 18).NumberFormat = "0.00%"
Next


''Création d'un dictionnaire des Mois commerciaux
'Set dico = CreateObject("Scripting.Dictionary")
'With dico
'    .Add "01", "M12"
'    .Add "02", "M1"
'    .Add "03", "M2"
'    .Add "04", "M3"
'    .Add "05", "M4"
'    .Add "06", "M5"
'    .Add "07", "M6"
'    .Add "08", "M7"
'    .Add "09", "M8"
'    .Add "10", "M9"
'    .Add "11", "M10"
'    .Add "12", "M11"
'End With
'
'
''Date
'For i = 10 To lastrow(2)
'Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'Next
''application du mois commercial à chaque case
'For i = 10 To lastrow(2)
'mois = Right(Left(Cells(i, 5), 5), 2)
'Cells(i, 1) = dico.Item(mois)
'Next


'Calculs des taux de commission de base regional

For i = 10 To lastrow(2)
    Select Case Cells(i, 16)
        Case Is < 100
            Cells(i, 19) = taux100
        Case Is > 100 / 100
            Cells(i, 19) = taux100max
    End Select
Cells(i, 19).NumberFormat = "0.0000%"

Next

'Calculs des taux de commission de base local '"QUADBenelux(CH06)"
For i = 10 To lastrow(2)
    If Cells(i, 2) = "QUADFrance(FR00)" Or Cells(i, 2) = "QUADBenelux(CH06)" Or Cells(i, 2) = "QUADDenmark(DK00)" Then
    
    Select Case Cells(i, 18)
        Case Is <= 100 / 100
            Cells(i, 20) = tx100local
            Cells(i, 20).NumberFormat = "0.00%"
        Case Is > 100 / 100
            
            Cells(i, 20) = tx100localmax
            Cells(i, 20).NumberFormat = "0.00%"
    End Select
    Else
    Cells(i, 20) = 0
Cells(i, 20).NumberFormat = "0.00%"
    End If
    
'If ActiveSheet.name <> "P. VAN ASSEM" Then
   ' If Cells(i, 2) = "QUADDenmark(DK00)" Then
     '   Cells(i, 20) = 0
    '    Cells(i, 20).NumberFormat = "0.0000%"
   '     Cells(i, 19) = 0
  '      Cells(i, 19).NumberFormat = "0.0000%"
 '   End If

'Else
    'If Cells(i, 2) = "QUADDenmark(DK00)" Then
    
    'Select Case Cells(i, 18)
        'Case Is <= 100
         '   Cells(i, 20) = tx100local
        '    Cells(i, 20).NumberFormat = "0.00%"
       ' Case Is > 100 / 100
      '      Cells(i, 20) = tx100localmax
     '       Cells(i, 20).NumberFormat = "0.00%"
    'End Select
    'End If
    
'End If

'calcul du commissionnemet en €
Cells(i, 21) = Cells(i, 19) * Cells(i, 14) + Cells(i, 20) * Cells(i, 14)
Cells(i, 21).NumberFormat = "0.00€"
Next


End Sub

Sub splitcommande_presales()

Dim palier, ponderation As Double
Dim obj, target, CA As Currency
Dim licence, maint, abon As Currency
Dim ligne As Integer

palier = 100

zone = Application.InputBox("Un palier de quel zone a été atteint? Entrer 1 ou local 2 pour régional")
If zone = 1 Then
    obj = Cells(5, 4)
ElseIf zone = 2 Then
    obj = Cells(3, 5)
    
Else: MsgBox "Valeur incorrecte, veuillez recommencez"
    Exit Sub
End If


Style = vbYesNo + vbCritical + vbDefaultButton2
Msg = "Confirmez-vous l'atteinte du palier des 100%?"


Response = MsgBox(Msg, Style)
If Response = vbYes Then    ' User chose Yes.
    palier = palier / 100
    target = obj * palier
    
    ligne = Application.InputBox("Quel est le numéro de ligne de la commande qui dépasse un palier?")
    
    
    If ligne = False Then
    
    Exit Sub
    End If
    
    'Si zone= 1 cumul local...
    If zone = 1 Then
    cumul = Cells(ligne - 1, 17)
    Else
    cumul = Cells(ligne - 1, 15)
    End If
    
    CA = Cells(ligne, 14)
    
    
    diff = target - cumul
    ponderation = diff / CA
    'licence = Cells(ligne, 7)
    'maint = Cells(ligne, 8)
    'abon = Cells(ligne, 9)
    
    'Insertion d'une ligne identique à celle à split'
    Rows(ligne + 1).Insert
    Rows(ligne).Copy
    Rows(ligne + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Rows(ligne).Font.ColorIndex = 14
    
    'modification des lignes de CA avant palier'
    
    Cells(ligne, 14) = ponderation * CA
    Cells(ligne, 14).Font.ColorIndex = 13
    
    'Cells(ligne, 8) = ponderation * maint
    'Cells(ligne, 8).Font.ColorIndex = 13
    
    'modification du montant de abo si non nul'
    'If Cells(ligne, 10) <> 0 Then
    '
    '    If Cells(ligne, 10) > 1 Then
    '
    '            Cells(ligne, 9) = Cells(ligne, 13) / Cells(5, 13)
    '        Else
    '            Cells(ligne, 9) = Cells(ligne, 13) / Cells(4, 13)
    '
    '    End If
    'Cells(ligne, 9).Font.ColorIndex = 13
    'End If
    
    'modification des lignes de CA après palier'
    Cells(ligne, 14).Offset(1, 0).Select
    Selection = CA * (1 - ponderation)
    Selection.Font.ColorIndex = 13
    
    'Cells(ligne, 8).Offset(1, 0).Select
    'Selection = maint * (1 - ponderation)
    'Selection.Font.ColorIndex = 13
    
    'Cells(ligne, 9).Offset(1, 0).Select
    'Selection = Cells(ligne, 9) * (1 - ponderation)
    'Selection.Font.ColorIndex = 13
    
    Call calculs_EMEA

Else    ' User chose No.
    Exit Sub
End If


'If palier = False Then
'    Exit Sub
'End If


End Sub

Sub dispatch_EMEA()

Dim start As Currency
Dim name As String
Dim wb As Workbook
Dim liste_name As Object
Dim vendeurs As Variant
Dim derligne As Integer

derligne = Cells(Rows.Count, 2).End(xlUp).Row

start = Application.InputBox("Ligne de démararrage ?")

Worksheets("Monthly Commissions").Activate
For i = start To lastrow(2)


    Range(Cells(i, 1), Cells(i, 18)).Copy
    For f = 2 To Worksheets.Count
        If Worksheets(f).name = "P. VAN ASSEM" Then
            Worksheets(f).Cells(lastrow(2, f) + 1, 1).PasteSpecial Paste:=xlPasteValues
        Else
            If Cells(i, 2) = "QUADDenmark(DK00)" Then
            
            Else
                Worksheets(f).Cells(lastrow(2, f) + 1, 1).PasteSpecial Paste:=xlPasteValues
            End If
        End If
        
    Next
Next

'recalcul pour chaque onglets
For f = 2 To Worksheets.Count
    Worksheets(f).Activate
    Call calculs_EMEA
Next

End Sub

