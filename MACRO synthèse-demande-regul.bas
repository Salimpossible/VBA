Attribute VB_Name = "Module1"
Sub import_demandes(filename, chemin)
Dim derligne As Integer
Dim dmd, wb As Workbook
Dim ws As Worksheet
Dim nom As Variant


Set wb = ThisWorkbook
Set ws = wb.Worksheets("DDE REGULS YTD 2021")
wb.Activate

'col = 2
'Call testderl(col)
'
'b = a

derlignelwb = Cells(Rows.Count, 2).End(xlUp).Row


Application.SendKeys ("{Enter}")
Workbooks.Open (chemin & "\" & filename)

Set dmd = Workbooks(filename)
dmd.Sheets(1).Activate
derligne = Cells(Rows.Count, 3).End(xlUp).Row
dercol = 14


'CC CV de la ligne info demandeur'
Range(Cells(5, 1), Cells(5, 6)).Select
dmd.Sheets(1).Range(Cells(5, 1), Cells(5, 6)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 1).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

'CC CV de la demande'
'demande.Activate
dmd.Sheets(1).Activate
Range(Cells(11, 1), Cells(derligne, dercol)).Copy
wb.Activate
ws.Cells(derlignelwb + 1, 7).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues

Cells(Cells(Rows.Count, 3).End(xlUp).Row, 1).Select

Range(Cells(Cells(Rows.Count, 3).End(xlUp).Row, 1), Cells(Cells(Rows.Count, 3).End(xlUp).Row, 6)).Copy

For i = Cells(Rows.Count, 3).End(xlUp).Row + 1 To Cells(Rows.Count, 15).End(xlUp).Row
Cells(i, 1).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues
Next

Workbooks(filename).Close saveChanges = False


End Sub

Sub loopfichiers(chemin)

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Dim chemin, fichier As Variant

Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFolder = oFSO.GetFolder(chemin)
    
    For Each oFile In oFolder.Files
    
        fichier = CStr(oFile.Name)
        
    Call import_demandes(fichier, chemin)
    
    Next oFile

End Sub


Sub selectfolder()

Dim doss As String

With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = -1 Then '-1 correspond à OK
        doss = .SelectedItems(1)
    End If
End With
If doss <> "" Then

Call loopfichiers(doss)

Else
    Exit Sub
End If
End Sub

