Attribute VB_Name = "Login"
'
' Login and init a session, interact with registry a bit
'
Option Explicit
Private username As String
Private password1 As String
Private salesforce As CSession     ' uses SFDC COM Toolkit module
Private Const no_user As String = "not logged in"
' list these reg keys in one place
Global Const VO_EN = "Volatile Environment"
Global Const SO_KEY = "Software\\Opensource"
Global Const SFDCUN = "SFDCUN"
Global Const SFDC = "SFDC"
Global Const SURL = "ServerURL"       ' for use in the option dialog
Global Const SETSURL = "SetServerURL"
Global Const GOAHEAD = "NoWarn"
Global Const NOLIMITS = "NoLimit"
Global Const CASERULE = "CaseRuleId"
Global Const LEADRULE = "LeadRuleId"
Global Const USE_NAMES = "UseNames"
Global Const USE_RELATED_CONTACT = "UseContactNames"
Global Const USE_RELATED_ACCOUNT = "UseAccountNames"
Global Const SPELL_USERNAME = "SpellOutUserNames"
Global Const NOT_FOUND = "#N/F"
Global Const SKIPHIDDEN = "SkipHiddenCells"  '** allows us to use a special update routine

Public Function ensureSalesforce()
  On Error Resume Next
  If (salesforce Is Nothing) Then ' fixed in 4.62, dont login each time thru here
    Set salesforce = New CSession
    If (validate("", "")) Then
      Call displayUserName(username)
    Else
      Call displayUserName(no_user)
    End If
  End If

' do i really need this, in 5.28 i simplify this away
' 5.08 test the session, may have been hung up
'  Dim st$: st = salesforce.GetServerTime
'  If st = "" Then ' 5.09 api error is not cleared when we just get server time
'    ' 5.08 if the session is invalid, need to ditch it
'    ' could be expired or intermitent network connection like wireless
'    Call GetSession ' we have the object, just need to refresh the session
'    st = salesforce.GetServerTime
'    If st = "" Then ' really didn't work
'        sfError "could not restore sforce session, login again" & vbCrLf & _
'        "Error msg :" & salesforce.GetErrorMessage
'        salesforce = Nothing ' dump it, will require a new login
'        End If
'  End If
  
  Set ensureSalesforce = salesforce
End Function

'Private Function GetSession()
'  GetSession = validate("", "")
'  If (GetSession) Then
'      Call displayUserName(username)
'  Else
'      Call displayUserName(no_user)
'  End If
'End Function

Public Function logout()
  password1 = ""
  username = ""
  Call SetKeyValue(HKEY_CURRENT_USER, VO_EN, SFDC, "", REG_SZ)
  Call SetKeyValue(HKEY_CURRENT_USER, VO_EN, SFDCUN, "", REG_SZ)
  Call displayUserName(no_user)
  On Error Resume Next
  Set salesforce = Nothing
End Function
Public Function getUserName()
  getUserName = username
End Function
Private Function validate(origUserName As String, password_in As String)
  On Error GoTo wild_error
  Dim success As Boolean
  ' get the server url from the registry if the user has placed one there
  ' by using the options box, or advanced on the login page
  ' otherwise use the default provided by the toolkit 5.15
  Dim defaulturl As String
  defaulturl = QueryValue(HKEY_CURRENT_USER, SO_KEY, SURL)
  
  loginForm.serverurl.value = IIf(defaulturl <> "", defaulturl, salesforce.serverurl)
  loginForm.serverurl.Enabled = False
  'Debug.Print "server url: " & loginForm.serverurl.value
  Do
      ' save original username to see if we're changing users
      username = origUserName
      If (username = "") Then
          username = QueryValue(HKEY_CURRENT_USER, VO_EN, SFDCUN)
      End If
      If (username = "") Then ' get it from another place..
            username = QueryValue(HKEY_CURRENT_USER, SO_KEY, SFDC)
      End If
      
      password1 = password_in
      If (password1 = "") Then
          password1 = QueryValue(HKEY_CURRENT_USER, VO_EN, SFDC)
      End If
      
      If (username = "" Or password1 = "") Then
          loginForm.username.value = username
          loginForm.password.value = password1
          If (username = "") Then
            loginForm.username.SetFocus
          Else
            loginForm.password.SetFocus
          End If
              
          loginForm.Show
          
          Application.StatusBar = "Authenticate user"
          username = loginForm.username.value
          password1 = loginForm.password.value
          
          ' if cancel login, clear password
          If (loginForm.cancel) Then
             Call logout ' clears out the session object also
             validate = False
             GoTo done
          End If
          
          ' rememember my name...
          Call SetKeyValue(HKEY_CURRENT_USER, SO_KEY, SFDC, username, REG_SZ)
    
      End If
      
      If (loginForm.serverurl.value <> "") Then
          salesforce.serverurl = loginForm.serverurl.value
      End If
      
      ' used to test how many times we try to re-login
      ' Debug.Print "calling dologin"
      
      success = salesforce.DoLogin(username, password1)

      If (Not success) Then ' if the password was incorrect, give a message
          salesforce.Msg_and_Fail_if_APIError ' pop a dialog with the bad news
          'MsgBox (" error number " & salesforce.wasFault)
          password1 = "" ' clear failed passwd
          Call SetKeyValue(HKEY_CURRENT_USER, VO_EN, SFDC, "", REG_SZ)
            
          validate = False ' loop around again
      End If
            
  Loop Until (success)
  If (QueryValue(HKEY_CURRENT_USER, VO_EN, SFDC) = "") Then
      Call SetKeyValue(HKEY_CURRENT_USER, VO_EN, SFDC, password1, REG_SZ)
  End If
  If (QueryValue(HKEY_CURRENT_USER, VO_EN, SFDCUN) = "") Then
      Call SetKeyValue(HKEY_CURRENT_USER, VO_EN, SFDCUN, username, REG_SZ)
  End If
  validate = True
GoTo done
wild_error:
  MsgBox "Salesforce: Login() " & vbCrLf & _
        "unknown error" & vbCrLf & _
        Error(), vbOKOnly + vbCritical
done:
  Application.StatusBar = False
End Function




