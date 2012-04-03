VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents sfApi As SForceSession4
Dim m_isValid As Boolean ' for login result
Dim UserNames As Scripting.Dictionary ' these dicts hold name-id pairs that we have looked up
Dim RecordTypes As Scripting.Dictionary ' makes for a much more readable display for these common fields
Dim Profiles As Scripting.Dictionary
Dim Roles As Scripting.Dictionary
Dim Groups As Scripting.Dictionary
Dim Contact As Scripting.Dictionary  ' caution, can cause performance issues
Dim Account As Scripting.Dictionary  ' caution, can cause performance issues
Dim redo_cmd As String ' track what we are trying to do

Private Sub Class_Initialize()
    Set sfApi = New SForceSession4
    Set UserNames = New Scripting.Dictionary
    Set RecordTypes = New Scripting.Dictionary
    Set Profiles = New Scripting.Dictionary
    Set Roles = New Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    Set Contact = New Scripting.Dictionary ' may become a large dict!!
    Set Account = New Scripting.Dictionary ' may become a large dict!!
End Sub

'
' given a users name or string name of a reference type
' return the id or any other case or
' if it's not a reference at all, ref_to will be nul, make that case
'  return the name_string, this is working in 5.46
'
Public Function ref_id(name_string As String, ref_to As String) As String
    Dim nam As Variant
    Dim qrs As QueryResultSet4, s As Object
    ref_id = name_string ' which may be an actual ID also... check for this
    
    Select Case Len(ref_id)
    Case 15, 18: ' length matches, and two adjacent numbers, looks like an ID 5.51
        ' kick out now or we may turn a real ID into a name string
        If (ref_id Like "*[0-9][0-9]*") Then Exit Function
    End Select
    
    On Error Resume Next: Select Case ref_to
    
    Case "User":
        If UserNames.Exists(name_string) Then
           ref_id = UserNames.Item(name_string):     Exit Function
        End If
        nam = Split(name_string, " ")
        Set qrs = query("select firstname,lastname,id from " & _
            "user where firstname = '" & nam(0) & "' and lastname = '" & nam(1) & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            UserNames.Add ref_id, name_string
            UserNames.Add name_string, ref_id
        Next s
    
    ' use of this type of reference should be controled by an option
    ' it could cause a serious performance problem on long queries
    ' and is not going to return unique strings if there are dups in the database
    ' should check for dups!!! and then return the ID passed in rather than guessing.
    ' as it does now TODO
    Case "Contact":
        ' 5.59 this is slow to open and close the registry each time, need to set a global
      '  If (Not QueryRegBool(USE_RELATED_CONTACT)) Then Exit Function ' dont map unless option is set for this... 5.59
       Exit Function ' before we can use this we need to check for duplicates on the retrun...
        If Contact.Exists(name_string) Then
           ref_id = Contact.Item(name_string):     Exit Function
        End If
        nam = Split(name_string, " ")
        Set qrs = query("select firstname,lastname,id from " & _
            "Contact where firstname = '" & nam(0) & "' and lastname = '" & nam(1) & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            Contact.Add ref_id, name_string
            Contact.Add name_string, ref_id
        Next s
        
    ' use of this type of reference should be controled by an option
    ' it could cause a serious performance problem on long queries
    ' and is not going to return unique strings if there are dups in the database
    ' should check for dups!!! and then return the ID passed in rather than guessing.
    ' as it does now TODO
    Case "Account":
        ' 5.59 this is slow to open and close the registry each time, need to set a global
     '   If (Not QueryRegBool(USE_RELATED_ACCOUNT)) Then Exit Function ' dont map unless option is set for this... 5.59
     Exit Function ' before we can use this we need to check for duplicates on the retrun...
        If Account.Exists(name_string) Then
           ref_id = Account.Item(name_string):     Exit Function
        End If
        Set qrs = query("select id,name from Account where Name = '" & name_string & "'")
        ' if we found more than one, just leave the "ID" !!
        ' TODO, count the returns
        ' TODO check for dups here, very important...
        ' don't want to map a specific string to a duplicate name...
        For Each s In qrs ' returns a random value !!
            ref_id = s.Item("Id").value
            Account.Add ref_id, name_string
            Account.Add name_string, ref_id
        Next s
        
        
    Case "RecordType":
        If RecordTypes.Exists(name_string) Then
           ref_id = RecordTypes.Item(name_string):     Exit Function
        End If
        Set qrs = query("select id,name from recordtype where Name = '" & name_string & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            RecordTypes.Add ref_id, name_string
            RecordTypes.Add name_string, ref_id
        Next s
    
    Case "Profile":
        If Profiles.Exists(name_string) Then
           ref_id = Profiles.Item(name_string):     Exit Function
        End If
        Set qrs = query("select id,name from profile where Name = '" & name_string & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            Profiles.Add ref_id, name_string
            Profiles.Add name_string, ref_id
        Next s
     
    Case "Group":
        If Groups.Exists(name_string) Then
           ref_id = Groups.Item(name_string):     Exit Function
        End If
        Set qrs = query("select id,name from group where Name = '" & name_string & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            Groups.Add ref_id, name_string
            Groups.Add name_string, ref_id
        Next s
        
        ' 5.51 none ? try again, may be a "UserOrGroupid" field, so recurse...
        If (ref_id = name_string) Then ref_id = ref_id(name_string, "User")
        
    Case "UserRole":
        If Roles.Exists(name_string) Then
           ref_id = Roles.Item(name_string):     Exit Function
        End If
        Set qrs = query("select id,name from userrole where Name = '" & name_string & "'")
        For Each s In qrs
            ref_id = s.Item("Id").value
            Roles.Add ref_id, name_string
            Roles.Add name_string, ref_id
        Next s
    
    ' and we arrive here for not a ref_to at all in 5.46
    Case Else ' 5.37 don't know how to map this type, so restore the value passed in
        ref_id = name_string ' assume it was correct and we got called by mistake.
        
    End Select

End Function
'
' lookup ID and return the string name and add the name to a dict
' works on user id's and record types currently, can be extended
' for roles, profiles, groups, etc.
'
Public Function ref_name(objectid As String, ref_to As String) As String
    Dim qrs As QueryResultSet4, s As Object
    If objectid = "" Then
        ref_name = "" ' we get here for Converted Account Id in leads which are not converted...
        Exit Function
    End If
    
    ref_name = objectid ' start with the id passed in, need to look it up or will return this
    
    ' note, sometimes group references are really userid's
    ' and othertimes organization id's are really userid's
    If Left(objectid, 3) = "005" Then ref_to = "User" '  was a user id passed in
    
    On Error Resume Next: Select Case ref_to
    
    Case "Contact":
        ' 5.59 this is slow to open and close the registry for each row, need to set a global TODO
        If (Not QueryRegBool(USE_RELATED_CONTACT)) Then Exit Function ' dont map unless option is set for this... 5.59
        
        If Contact.Exists(objectid) Then
            ref_name = Contact.Item(objectid):   Exit Function
        End If
        Set qrs = query("select firstname,lastname,id from Contact where Id = '" & objectid & "' ")
        For Each s In qrs
            Debug.Assert objectid = s.Item("Id").value ' better be true
            ref_name = sfNameToVal(s, "FirstName") & " " & sfNameToVal(s, "LastName")
            Contact.Add s.Item("Id").value, ref_name
            Contact.Add ref_name, s.Item("Id").value
        Next s
        
    Case "Account":
        ' 5.59 this is slow to open and close the registry for each row, need to set a global TODO
        If (Not QueryRegBool(USE_RELATED_ACCOUNT)) Then Exit Function ' dont map unless option is set for this... 5.59
        
        If Account.Exists(objectid) Then
            ref_name = Account.Item(objectid):   Exit Function
        End If
        Set qrs = query("select id,name from Account where Id = '" & objectid & "' ")
        ' what about DUPS, we may insert duplicate records here....
        ' which will crash with "key already exists"
        For Each s In qrs
            ref_name = s.Item("Name").value
            Account.Add s.Item("Id").value, ref_name
            Account.Add ref_name, s.Item("Id").value
        Next s
     
    Case "User":
        ' would be nice to look up and return the org info here...
        ' since this is located int the session object we can pull it
        If UserNames.Exists(objectid) Then
            ref_name = UserNames.Item(objectid):   Exit Function
        End If
        Set qrs = query("select firstname,lastname,id from user where Id = '" & objectid & "' ")
        For Each s In qrs
            Debug.Assert objectid = s.Item("Id").value ' better be true
            ref_name = sfNameToVal(s, "FirstName") & " " & sfNameToVal(s, "LastName")
            UserNames.Add s.Item("Id").value, ref_name
            UserNames.Add ref_name, s.Item("Id").value
        Next s
 
    Case "RecordType":
        If RecordTypes.Exists(objectid) Then
            ref_name = RecordTypes.Item(objectid):   Exit Function
        End If
        Set qrs = query("select id,name from recordtype where id = '" & objectid & "'")
        For Each s In qrs
            ref_name = s.Item("Name").value
            RecordTypes.Add s.Item("Id").value, ref_name
            RecordTypes.Add ref_name, s.Item("Id").value
        Next s

    Case "Profile":
        If Profiles.Exists(objectid) Then
            ref_name = Profiles.Item(objectid):   Exit Function
        End If
        Set qrs = query("select id,name from profile where id = '" & objectid & "'")
        For Each s In qrs
            ref_name = s.Item("Name").value
            Profiles.Add s.Item("Id").value, ref_name
            Profiles.Add ref_name, s.Item("Id").value
        Next s
    
    Case "Group":
        If Groups.Exists(objectid) Then
            ref_name = Groups.Item(objectid):   Exit Function
        End If
        Set qrs = query("select id,name from group where id = '" & objectid & "'")
        For Each s In qrs
            ref_name = s.Item("Name").value
            
            ' 5.40 sometimes the group name is empty stash the id instead of an empty name
            If (s.Item("Name").value = "") Then ref_name = s.Item("Id").value
            
            Groups.Add s.Item("Id").value, ref_name
            Groups.Add ref_name, s.Item("Id").value
        Next s
        
    Case "UserRole":
        If Roles.Exists(objectid) Then
            ref_name = Roles.Item(objectid):   Exit Function
        End If
        Set qrs = query("select id,name from userrole where id = '" & objectid & "'")
        For Each s In qrs
            ref_name = s.Item("Name").value
            Roles.Add s.Item("Id").value, ref_name
            Roles.Add ref_name, s.Item("Id").value
        Next s
    
     Case "Organization": ' org contains a user name OR the ultimate org id for your org..
        If FixID(objectid) = FixID(sfApi.OrganizationId) Then
            ref_name = sfApi.OrganizationName
            UserNames.Add s.Item("Id").value, ref_name
            UserNames.Add ref_name, s.Item("Id").value
        End If
     
     Case Else
        ref_name = objectid ' too small, throw it back
        
    End Select
    
End Function

Public Property Let serverurl(url As String)
    sfApi.SetServerUrl (url)
End Property
Public Property Get serverurl() As String
    serverurl = sfApi.serverurl
End Property
Public Function IsValid() As Boolean
    IsValid = m_isValid
End Function
'
' used to pass instructions on case and lead create
' can cause assignment rules to fire
'
Public Sub SetSoapHeader(section As String, name As String, value As String)
    sfApi.SetSoapHeader section, name, value
    
End Sub

'2004-03-10T21:22:26.000Z
'4/12/2004 9:13:17 PM
Public Function GetServerTime() As String
    Dim ct As Date: ct = sfApi.CurrentServerTime
    GetServerTime = Format$(ct, "yyyy-mm-ddTHH:MM:SS.000Z")
End Function

Public Function DoLogin(uname As String, pword As String) As Boolean
    m_isValid = sfApi.Login(uname, pword)
'   just to be sure it is working, do a get time
'    Dim stime
'    On Error Resume Next
'    stime = sfApi.CurrentServerTime ' make one call to see that the login is ok
'    If (sfApi.Error > 0) Then m_isValid = False ' 6.14
    DoLogin = m_isValid
End Function

Public Function CreateEntity(entityType As String) As SObject4
    Set CreateEntity = sfApi.CreateObject(entityType)
End Function

Public Function DoCreate(sa() As SObject4) As Boolean

On Error Resume Next
   
    Call sfApi.Create(sa, False)
    
    If sfApi.Error > 0 Then
        DoCreate = False
    Else
    
        DoCreate = True
            
        If sfApi.Error > 0 Then
            DoCreate = False
        Else
            DoCreate = True
        
        End If
    
    End If
    
End Function
Public Function DoDelete(idlist() As String, entityType As String) As Boolean

    On Error Resume Next
    Dim qr As QueryResultSet4
    Set qr = sfApi.Retrieve("Id", entityType, idlist, False)
    If sfApi.Error > 0 Then
        DoDelete = False
        Exit Function
        End If
        
    Dim sa() As SObject4
    Dim s As SObject4
    '  If (m_numObjects = 0) Then m_numObjects = 2000
    ReDim sa(2000)
    Dim x As Integer
    x = 0
    For Each s In qr
    Set sa(x) = s
        x = x + 1
    Next s
    ReDim Preserve sa(x - 1)
    
    sfApi.Delete sa, False
    
    If sfApi.Error > 0 Then
        DoDelete = False
    Else
    
        DoDelete = True
        
        sfApi.Refresh sa, False
        
        If sfApi.Error > 0 Then
            DoDelete = False
        Else
            DoDelete = True
        
        End If
    
    End If
    
End Function

Public Function query(q) As QueryResultSet4
    Set query = sfApi.query(q, False) ' not async
    ' if this fails due to an invalid session, we can retry ?
    ' If (sfApi.Error = 5084) Then   ' still working on this
    '    Debug.Print "session timeout returned from query"
        ' get a new session
        ' try the query again...
   ' End If
End Function
Public Function wasFault()
    wasFault = sfApi.Error
End Function

Public Function DoUpdate(sa() As SObject4) As Boolean

On Error Resume Next

    sfApi.update sa, False
    
    If sfApi.Error > 0 Then
        DoUpdate = False
    Else
    
        DoUpdate = True
        
        If sfApi.Error > 0 Then
            DoUpdate = False
        Else
            DoUpdate = True
        
        End If
    
    End If
    
End Function
'
' what if there are duplicates in the idlist ?
'
Public Function DoFullGet(idlist() As String, entityType As String) As Object

     Dim qr As QueryResultSet4
     
     'm_numObjects = UBound(idlist) + 1
     
     Set qr = sfApi.Retrieve("*", entityType, idlist, False)
    
     Set DoFullGet = ProcessQueryResults(qr)
     
End Function
'
' passing a limited sels list keeps this a fast routine.
'
Public Function Retrieve(sels As String, idlist() As String, entityType As String) As Object
    ' make sure id list has no blanks
    Dim i%: For i = 0 To UBound(idlist)
      Debug.Assert idlist(i) <> "" ' bad things ...
    Next i
    Set Retrieve = sfApi.Retrieve(sels, entityType, idlist, False)
End Function
'
' load it up in a dictionary of object by id
' NOTE, the query result may be more than 2000 long, in which case
' we probably would miss something, ouch...
' TODO find out if this is happening ?
'
Public Function ProcessQueryResults(ByVal qr As QueryResultSet4) As Object

    Dim s As SObject4
    Dim sd As New Scripting.Dictionary
    Dim sa() As SObject4
    Dim x As Integer
    
  '  If (m_numObjects = 0) Then m_numObjects = 2000
    ReDim sa(2000)
    
    For Each s In qr
        
        Set sa(x) = s
        
        x = x + 1
    
    Next s
    
    ' debug.assert x <2000 ' or we truncated the query... TODO test this
    
    If x = 0 Then
        'did not get everything i hoped for
        Set ProcessQueryResults = Nothing
        Exit Function
    End If
    ReDim Preserve sa(x - 1)
            
    For x = 0 To UBound(sa)
        ' avoid adding dups to this collection
        If Not sd.Exists(sa(x).Item("Id").value) Then
          sd.Add sa(x).Item("Id").value, sa(x)
        End If
    Next x
    
    Set ProcessQueryResults = sd

End Function

Public Function GetUserId() As String
    GetUserId = sfApi.userid
End Function

Public Function GetError() As Integer
    GetError = sfApi.Error
End Function

Public Function GetErrorMessage() As String
    GetErrorMessage = sfApi.ErrorMessage
End Function

Function EntityNames()
    EntityNames = sfApi.EntityNames
    If sfApi.Error > 0 Then
        MsgBox sfApi.ErrorMessage
    End If
End Function

Private Function DisplayStatusText(label As String, Optional value As String = vbCrLf)
    Dim str As String
    str = str & label & value & vbCrLf
    MsgBox str
End Function
'
' renamed to preserve my sanity, (opps too late...)
' if we check it for an API error and there is one,
'   we msg user and return false
' otherwise no error, the check routine returns true
'
Function Msg_and_Fail_if_APIError() As Boolean
    If sfApi.Error > 0 Then
        Dim oops$: oops = sfApi.ErrorMessage & vbCrLf & "ExceptionCode : " & sfApi.Error
        ' if it's an invalid session due to timeout,
        ' just log them out, and ask them to re-run the command that was just run...
        If (sfApi.Error = 5084) Then
            sfLogout
            oops = oops & vbCrLf & vbCrLf & _
                "Session Timeout: Please re-run the command >" & redo_cmd
        End If
        If (sfApi.Error = 1) Then
            sfLogout
            oops = oops & vbCrLf & vbCrLf & _
                "Could not establish connection to API service" ' 6.14
        End If
        DisplayStatusText "Error Generated by request::", oops
        Msg_and_Fail_if_APIError = False
        Exit Function
    End If
    Msg_and_Fail_if_APIError = True
End Function

Function Search(sosl As String) As Object
    Dim qr As QueryResultSet4
    Set qr = sfApi.Search(sosl, False)
    Set Search = ProcessQueryResults(qr)
End Function
