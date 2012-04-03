VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Step2b 
   Caption         =   "Sforce Table Query Wizard - Step 2 bis of 3"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   OleObjectBlob   =   "Step2b.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Step2b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call Step2bField_KeyPress(KeyAscii)
End Sub

' others keys are possible to catch here
Private Sub Step2bField_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Debug.Print "key ascii is " & KeyAscii
    Select Case KeyAscii
      Case 1 ' select all CTRL-A
        Call UserForm_Initialize
      Case 27 ' escape
        Me.Hide
    End Select
End Sub


Private Sub UserForm_Activate()
    'Call selectAll
    ' better to remember from last time we visited here?
End Sub

Public Function select_all_fields()
    ' default to all items selected
    For i = 0 To Me.Step2bField.ListCount - 1
        Me.Step2bField.Selected(i) = True
    Next i
End Function

Private Sub UserForm_Initialize()
    Call select_all_fields
End Sub

Private Sub btnBack_Click()
    Step2b.Hide
    describeBox.Show
End Sub
'
Private Sub btnCancel_Click()
    Step2b.Hide
End Sub

' go to the next dialog
Private Sub btnNext_Click()
    Call ensure_ID_select(Me.Step2bField)
    Step2b.Hide
    s_force.sfDescribe_wiz_draw ' draw the fields that were selected
    queryWiz.Show     ' then go to final step of wizard
End Sub

' given a label, see if it has been selected in this list
Public Function is_wiz_selected(label)
    For i = 0 To Me.Step2bField.ListCount - 1
        If (Me.Step2bField.List(i) = label) Then
            If (Me.Step2bField.Selected(i)) Then
                is_wiz_selected = True
                Exit Function
            End If
        End If
    Next i
    is_wiz_selected = False
End Function

