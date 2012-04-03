VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginForm 
   Caption         =   "salesforce.com login"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   OleObjectBlob   =   "loginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "loginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cancel As Boolean

Private Sub advanced_Click()
'serverurl.Locked = False
serverurl.Enabled = True
serverurl.SetFocus
End Sub

Private Sub cancelButton_Click()
    cancel = True
    Hide
End Sub

Private Sub CommandButton1_Click()
    cancel = False
    Hide
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

