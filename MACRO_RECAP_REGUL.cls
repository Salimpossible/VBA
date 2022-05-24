VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub import_demandes(filename, chemin)
Dim derligne As Integer
Dim dmd, wb As Workbook
Dim ws As Worksheet
Dim nom As Variant


Set wb = ThisWorkbook
Set ws = wb.Worksheets("DDE REGULS YTD 2021")
wb.Activate
derlignelwb = lastrow(2)



Application.SendKeys ("{Enter}")
Workbooks.Open (chemin & "\" & filename)

Set dmd = Workbooks(filename)
dmd.Sheets(1).Activate
derligne = lastrow(3)
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

Cells(lastrow(3), 1).Select

Range(Cells(lastrow(3), 1), Cells(lastrow(3), 6)).Copy

For i = lastrow(3) + 1 To lastrow(15)
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
