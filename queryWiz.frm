VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} queryWiz 
   Caption         =   "Sforce Table Query Wizard - Step 3 of 3"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   OleObjectBlob   =   "queryWiz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "queryWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim g_start As Range ' where the range starts on the sheet, could be "A1" or else where

Private Sub addtoquery_Click()
    ' load it into the cells just to the right of the current query end
    ' assume that the current selection is the end of current query
    ' or is exactly on the "g_start" cell, i.e: no query added yet
    '
    ' if none please select a field value
    If (field.SelText = "--Select One--") Then
      MsgBox "Please select a Field"
      GoTo done
    End If
    
    ' TODO validate the values?
    ' really should do field type checking here
    ' or it implies by placing it in the spreadsheet that it will work in SOQL

    Selection.Offset(0, 1).value = field.SelText
    Selection.Offset(0, 2).value = operator.SelText
    Selection.Offset(0, 3).value = values.Text
    
    Call AddClause((Selection.Column - g_start.Column) / 3) ' add to list object
    
    Selection.Offset(0, 3).Select ' move selection , right shift by 3
done:
End Sub

'
' clear one clause, first clear query area, then remove current selected list item
' finaly re-draw query using remaining list items
' 5.42
Private Sub clearOne_Click()
    Call clear_queryarea
    ' remove the current selected value from the list in the list widget
    With Me.queryClausesList
        If .ListIndex = -1 Then GoTo done
        .RemoveItem (.ListIndex)
    
        ' using data found in the list box, load the query area
        Dim i%:  For i = 0 To .ListCount - 1
            Selection.Offset(0, 1).Select: Selection.value = .List(i, 1)
            Selection.Offset(0, 1).Select: Selection.value = .List(i, 2)
            Selection.Offset(0, 1).Select: Selection.value = .List(i, 3)
        Next i
    End With

done:
End Sub
Private Sub UserForm_Activate()
    Dim used As Range ' if the worksheet is empty, move the selection to the top left
    Set used = Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell).Address)
    If used.Count = 1 Then Range("A1").Select
    
    s_force.queryWizard_init ' loads picklists with field names and operators
    
    ' populate the list object and
    ' move the selection to the end of the current existing query
    Call PopulateClausesList
End Sub
'
' using the data found in the query area, load the listbox with items.
' 5.42
Sub PopulateClausesList()
    Dim clauseNum As Integer
    
    Set g_start = ActiveCell.CurrentRegion.Cells(1, 1)
    g_start.Select
    Me.queryClausesList.Clear
    clauseNum = -1
    Do While (Selection.Offset(0, 1).value <> "")
      clauseNum = clauseNum + 1
      Call AddClause(clauseNum)
      Selection.Offset(0, 3).Select
    Loop
    Me.clearOne.Enabled = False ' start with none selected, button off
End Sub
Sub AddClause(clauseNum As Integer) ' 5.42
    With Me.queryClausesList
        .AddItem clauseNum, clauseNum
        .List(clauseNum, 1) = Selection.Offset(0, 1)
        .List(clauseNum, 2) = Selection.Offset(0, 2)
        .List(clauseNum, 3) = Selection.Offset(0, 3)
    End With
End Sub
'
' clear all name,op,value sequences from the query row
' leave the table name only and selected need to use clear instead
' of delete incase other tables exist on this sheet
'
Private Sub clear_queryarea()
    g_start.Select
    Selection.Offset(0, 1).Select  ' move one to the right to begin at the first field
    While (Selection.Offset(0, 0).value <> "")
        ' clear three cells only
        Range(Selection.Cells(1, 1), Selection.Cells(1, 3)).Clear
        Selection.Offset(0, 3).Select ' move 3 right
    Wend
    g_start.Select
End Sub

Private Sub clearAll_Click() ' 5.42
    Call clear_queryarea: Me.queryClausesList.Clear
    Me.clearOne.Enabled = False
End Sub
Private Sub queryClausesList_Change() ' 5.42
    Me.clearOne.Enabled = Not IsNull(Me.queryClausesList.value)
End Sub
Private Sub UserForm_Initialize()
    ' nothing here, do the work in _Activate
End Sub
Private Sub back_Click()
    queryWiz.Hide
End Sub
Private Sub run_Click()
    queryWiz.Hide: s_force.sfQuery
End Sub
