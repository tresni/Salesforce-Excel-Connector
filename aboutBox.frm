VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aboutBox 
   Caption         =   "About the sforce Connector"
   ClientHeight    =   3165
   ClientLeft      =   195
   ClientTop       =   360
   ClientWidth     =   5325
   OleObjectBlob   =   "aboutBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Label2_Click()

End Sub

Private Sub CommandButton2_Click()

' open using the user default browser, thanks to Christophe Humbert
Set WshShell = CreateObject("WScript.Shell")
WshShell.run "rundll32 url.dll,FileProtocolHandler " & _
    "http://code.google.com/p/excel-connector/w/list"
WshShell = ""

done:
aboutBox.Hide
End Sub

Private Sub CommandButton3_Click()
aboutBox.Hide
End Sub

Private Sub UserForm_Initialize()
 version.Caption = AutoExec.ver_str()
End Sub
