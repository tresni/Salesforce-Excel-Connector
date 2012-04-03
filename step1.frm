VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} step1 
   Caption         =   "Sforce Table Query Wizard - Step 1 of 3"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   OleObjectBlob   =   "step1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub cancel_Click()
step1.Hide
End Sub
Private Sub gonext_Click()
step1.Hide
describeBox.Show
End Sub
Private Sub location_Change()
  ActiveSheet.Range(location.Text).Select ' keep the selection in sync with user
End Sub
Private Sub UserForm_Activate()
  ' if the worksheet is empty, move the selection to the top left
  Dim used As Range
  On Error GoTo nosheet
  Set used = Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell).Address)
  If used.Count = 1 Then Range("A1").Select
  
  ' else move it to the upper left of the current region
  ActiveCell.CurrentRegion.Cells(1, 1).Select
  
  location.Text = Selection.Address
  GoTo done
nosheet:
  MsgBox "Oops, Could not find an active Worksheet"
done:
End Sub
Private Sub UserForm_Initialize()
  ' grab the current selection and put it in the location object
  location.Text = Selection.Address
End Sub
