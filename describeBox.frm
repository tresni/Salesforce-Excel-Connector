VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} describeBox 
   Caption         =   "Sforce Table Query Wizard - Step 2 of 3"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4185
   OleObjectBlob   =   "describeBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "describeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
' This is step 2A of the Wizard
' Wizard type functionality for quick start with new uesrs
' thanks go to Adam G. for the kick start on this module
'
Private Sub describe_Click()
  ' grab the desired location
  ActiveSheet.Range(step1.location.Text).Select ' fails on a second worksheet
  tVal = describeBox.describeResultBox.Text
  If tVal = "" Then
    MsgBox "please select an Entity from the Pick-List, or Cancel"
    GoTo done
  End If
  
  Selection.Cells(1, 1).value = tVal
  describeBox.Hide
  
  ' ask if they want to overwrite existing column names
  If (Selection.Cells(2, 1).value <> "") Then
    Dim vRet: vRet = MsgBox("Overwrite existing Column names ?", _
      vbApplicationModal + vbYesNo + vbExclamation + vbDefaultButton1, _
         "-- Overwrite Existing? --")
    If Not (vRet = vbYes) Then GoTo almostdone
  End If
  
  s_force.sfDescribe_wiz ' loads the step2B list
  Call Step2b.select_all_fields
  
almostdone:
  'queryWiz.Show 'comment out CHM 04/2005
  Step2b.Show
done:
End Sub

Private Sub describeResultBox_DblClick(ByVal cancel As MSForms.ReturnBoolean)
  describe_Click ' short cut for power users
End Sub

'Private Sub describeSObject_Click()
'  describe_Click
'End Sub

Private Sub cancel_Click()
  describeBox.Hide
End Sub

Private Sub describeSObject_Click()
describe_Click
End Sub

Private Sub goback_Click()
 describeBox.Hide
 step1.Show
End Sub

Private Sub UserForm_Activate()
 ' if the current cell is in our list, pop that into the select box
 Dim cur$
 cur = Selection.Cells(1, 1).value
 For i = 0 To describeBox.describeResultBox.ListCount - 1
  If (describeBox.describeResultBox.List(i, 0) = cur) Then
    'Debug.Print "here " & cur
    describeBox.describeResultBox.ListIndex = i
    GoTo done
    End If
 Next i
done:
End Sub

