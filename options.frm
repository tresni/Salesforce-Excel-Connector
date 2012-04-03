VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} options 
   Caption         =   "Sforce Connector Options"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   OleObjectBlob   =   "options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'
' each time this dialog comes up , set the flags per the reg
' added in v5.15
'
Private Sub UserForm_Activate()
  setserver.value = QueryRegBool(SETSURL)
  serverurl.Text = QueryValue(HKEY_CURRENT_USER, SO_KEY, SURL)
  
  If IsNull(setserver.value) Then setserver.value = False
  serverurl.Enabled = setserver.value ' dim the text box if needed
  
  nowarn.value = QueryRegBool(GOAHEAD)
  nolimit.value = QueryRegBool(NOLIMITS)
  caseassign.value = QueryRegBool(CASERULE)
  leadassign.value = QueryRegBool(LEADRULE)
  usenames.value = QueryRegBool(USE_NAMES)
  lookup_contacts.value = QueryRegBool(USE_RELATED_CONTACT)
  lookup_Accounts.value = QueryRegBool(USE_RELATED_ACCOUNT)
  skiphide.value = QueryRegBool(SKIPHIDDEN)
End Sub
'
' write our options to the registry
' added in v5.15
'
Private Sub done_Click()
  Dim setme As String: setme = "" ' default
  
  If setserver.value Then setme = serverurl.Text
  ' else it will be cleared out
  
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, SURL, setme, REG_SZ)
  serverurl.Text = setme
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, GOAHEAD, nowarn.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, SETSURL, setserver.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, NOLIMITS, nolimit.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, CASERULE, caseassign.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, LEADRULE, leadassign.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, USE_NAMES, usenames.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, USE_RELATED_CONTACT, lookup_contacts.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, USE_RELATED_ACCOUNT, lookup_Accounts.value, REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, SKIPHIDDEN, skiphide.value, REG_SZ)

  options.Hide
End Sub

Private Sub setserver_Click()
  serverurl.Enabled = setserver.value
End Sub
  
Private Sub cancel_Click()
options.Hide
End Sub
Private Sub CommandButton1_Click() ' page 2
options.Hide
End Sub

Private Sub CommandButton2_Click() ' page 2
Call done_Click
End Sub

Private Sub CommandButton3_Click() ' page 3
options.Hide
End Sub

Private Sub CommandButton4_Click()
done_Click
options.Hide
End Sub
Private Sub CommandButton5_Click() ' related names
options.Hide
End Sub
Private Sub CommandButton6_Click() ' related names apply
Call done_Click
End Sub
Private Sub CommandButton8_Click() ' filter apply
Call done_Click
End Sub

Private Sub CommandButton7_Click()
options.Hide
End Sub

