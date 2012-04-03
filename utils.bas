Attribute VB_Name = "utils"
Option Explicit
'
' several helper functions follow
' mostly called from s_force module
' code which does not use globals or SObject4s is found here
'
'

' fixes all the ID's in a column (usually obtained from a sf.com report or a backup zip file) in-place.
'
' thanks to Scot S.
'
Sub fixidsel() ' operate on the selection
   Dim IDcol As Range: Set IDcol = ActiveCell.CurrentRegion.Columns(ActiveCell.Column)
   Dim c: For Each c In Intersect(IDcol, Selection)
        If Len(c) = 15 Then c.value = FixID(c.value)
   Next c
End Sub
Sub fixidcol() ' on the entire column
   Dim IDcol As Range: Set IDcol = ActiveCell.CurrentRegion.Columns(ActiveCell.Column)
   Dim c: For Each c In IDcol.Cells
        If Len(c) = 15 Then c.value = FixID(c.value)
   Next c
End Sub


'
' Converts a 15 character ID to an 18 character, case-insensitive one ...
' got this one from sforce community
' thanks go to Scot Stoney
'
Function FixID(InID As String) As String
If Len(InID) = 18 Then
  FixID = InID
  Exit Function
  End If
Dim InChars As String, InI As Integer, InUpper As String
Dim InCnt As Integer
InChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ012345"
InUpper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

InCnt = 0
For InI = 15 To 1 Step -1
   InCnt = 2 * InCnt + Sgn(InStr(1, InUpper, Mid(InID, InI, 1), vbBinaryCompare))
   If InI Mod 5 = 1 Then
       FixID = Mid(InChars, InCnt + 1, 1) + FixID
       InCnt = 0
       End If
    Next InI
    FixID = InID + FixID
End Function

' this is a very handy function to call from a cell in a worksheet
Function sfaccount(account_name)
 ' before we call, pull off common suffixes which may mess up our search
  Dim tmp$: tmp = Replace(account_name, "LLC", "")
  tmp = Replace(tmp, "Company", "")
  tmp = Replace(tmp, "Corporation", "")
  tmp = Replace(tmp, "Corp", "")
  tmp = Replace(tmp, "Inc", "")
  tmp = Replace(tmp, ".", "")
  tmp = Replace(tmp, ",", "")
  sfaccount = sfSearch("Account", account_name, "NAME")
  If (sfaccount = "#N/F") Then ' try exact match first, then fuzzy
    sfaccount = sfSearch("Account", tmp, "NAME")
  End If
End Function

Function sfcontact(firstname, Optional lastname)
  sfcontact = "#N/A"
  If IsMissing(lastname) Then
    Dim tmp() As String: tmp = Split(firstname, " ", 2)
    firstname = tmp(0)
    lastname = tmp(1)
  End If
  sfcontact = sfSearch("Contact", firstname & " " & lastname, "NAME")
End Function
Function sfemail(email_string, Optional findfirst)
  sfemail = sfSearch("Contact", email_string, "EMAIL", findfirst)
End Function

Function sflookup_all(table, cel)
  sflookup_all = sfSearch(table, cel, "all")
End Function
Function sflookup(table, cel) ' as long as you are passing the "name" ok
  sflookup = soql_table(table, " name  = '" & cel & "'")
End Function

Function soql_opp(cel) ' cel has the query clause ie:  " name like 'foo' "
 soql_opp = soql_table("Opportunity", cel)
End Function
Function soql_account(cel)
 soql_account = soql_table("Account", cel)
End Function
'
' if no table name is given, try to guess just using the header row
'
Function guess_table(table As Range)
  Dim c: For Each c In table.Rows(1).Cells
    Dim tmp As String: tmp = LCase(c.value)
    If (tmp = "opportunity id") Then guess_table = "opportunity"
    If (tmp = "account id") Then guess_table = "account"
    If (tmp = "contact id") Then guess_table = "contact"
    If (tmp = "case id") Then guess_table = "case"
    
    ' look for "Order :ID"
    If (Right(tmp, 4) = ": id") Then
        ' pull off the right side, then guess a custom table
        ' why do we need to put 's' on the end...
        ' name is the plural but the label is singluar!
        ' this wont always work...
        guess_table = Left(c.value, Len(c.value) - 4) & "s__c"
    End If
    If guess_table <> "" Then GoTo done
    Next c
done:
End Function

'
' helper funcition cause I could not figure out how to do this quickly in the worksheet...
'
Function QuarterNum(enter_date)
 QuarterNum = DatePart("q", enter_date)
End Function
'
' another helper...
'
Function NextQuarter(enter_date)
    ' return the first date of the next quarter
     Dim QuarterNum: QuarterNum = DatePart("q", enter_date)
     Dim tmp: tmp = DatePart("q", enter_date)
     Do While (QuarterNum = tmp)
        enter_date = enter_date + 1
        tmp = DatePart("q", enter_date)
        Loop
    NextQuarter = enter_date
End Function
'
' these are not good to pass to salesforce search without
' an escaping char
'
' TODO this only deals with some of these chars..
' & | ! ( ) { } [ ] ^ " ~ * ? : \ ' -
'
' asking to be optimized...
'
Function SFDC_escape(s As String) As String
Dim InI As Integer
' InChars = "&|!()[]^""~*?:'"
For InI = 1 To Len(s) Step 1
 ' Debug.Print Mid(s, InI, 1): Debug.Print Asc(Mid(s, InI, 1))
  
  Select Case Asc(Mid(s, InI, 1))
  Case 33 ' this is the ! char
    s = Left(s, InI - 1) & Chr(92) & Chr(33) & Right(s, Len(s) - InI)
    InI = InI + 1
  Case 38 ' this is the & char
    s = Left(s, InI - 1) & Chr(92) & Chr(38) & Right(s, Len(s) - InI)
    InI = InI + 1
  Case 45 ' this is the '-'
    s = Left(s, InI - 1) & Chr(92) & Chr(45) & Right(s, Len(s) - InI)
    InI = InI + 1
  Case 43 ' which ?
    s = Left(s, InI - 1) & Chr(92) & Chr(43) & Right(s, Len(s) - InI)
    InI = InI + 1
  Case 39 ' tic escape the tic is not working on 12/15/04
  ' revisit this when the regression is fixed
    s = Left(s, InI - 1) & Chr(92) & Chr(39) & Right(s, Len(s) - InI)
    InI = InI + 1
  End Select
  
Next InI

SFDC_escape = Trim(s) ' 5.65
End Function
'
' slightly different than above for query strings
'
Function SFDC_escape_q(s As String) As String
Dim InI As Integer
'  "&|!()[]^""~*?:'" should really deal with all of these, just lazy i guess
For InI = 1 To Len(s) Step 1
 ' Debug.Print Mid(s, InI, 1): Debug.Print Asc(Mid(s, InI, 1))
  Select Case Asc(Mid(s, InI, 1))
  Case 39 ' this is the tick ->'<-
    s = Left(s, InI - 1) & Chr(92) & Chr(39) & Right(s, Len(s) - InI)
    InI = InI + 1
  End Select
  
Next InI

SFDC_escape_q = Trim(s)
End Function
'
' look at the field type and variant type, cast the into a return value
'
Public Function toVBtype(value As Range, field)
    
    Select Case field.Type
    Case "i4":
        Dim i As Integer
        i = Int(Val(value))
        toVBtype = i
        
    ' 6.01
    ' re-do percent and currency for v3.0 toolkit
    ' percent is now it's own type
    Case "percent"
         toVBtype = Val(value) ' normal case
        ' special case, deal with excel cells formated as percentages,
        ' in general salesforce expects percentages to be doubles > 1.0
        ' and excel stores them as  numbers less than 1, ie: 55% == "0.55" in excel
        If Right(value.NumberFormat, 1) = "%" Then toVBtype = toVBtype * 100
    
    Case "double", "currency"
        ' val() does not use i18n conventions, use CDbl instead, 6.08
        toVBtype = CDbl(value)  ' normal case
        ' 6.01 truncate to the number of digits, Field3 likes it's numbers formated
        If (field.Scale = 0) Then
            toVBtype = Int(toVBtype)
        Else ' If (field.Scale > 0) Then
            Dim z: z = InStr(value, Application.International(xlDecimalSeparator))
            If (z > 0) Then  ' need to remove any extra decimal places
                toVBtype = CDbl(Left(value, z + field.Scale))
            End If
        End If
        
    Case "datetime", "date"
        Dim unInitVariant As Variant
        toVBtype = IIf(IsEmpty(value), unInitVariant, value) ' 5.13 handle empty dates
    
    Case "boolean":
        toVBtype = value
    
    Case Else ' all other types (so far),  work with this "string" type
      toVBtype = "" & value
    
    End Select
End Function

Function sfError(msg As String)
    ' put up a simple error, avoid using a new error form
    Call MsgBox(msg, vbOKOnly + vbExclamation, "Salesforce Excel Add-In")
    
End Function

'
' anyone have any ideas how to make this localized ??
' i found this which helped me (5.66)
' http://www.oaltd.co.uk/ExcelProgRef/Ch22/ProgRefCh22.htm
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnovba00/html/InternationalIssues.asp
' which pointed me to Application.International(xlDateOrder)
'
Function typeToFormat(sfType As String)
    typeToFormat = "General" ' default
    Select Case sfType
    
    Case "date", "datetime" ' re-written for 5.66
        Select Case Application.International(xlDateOrder)
        Case 0      'month-day-year
            typeToFormat = "m/d/yy"
        Case 1      'day-month-year
            typeToFormat = "d/m/yy"
        Case 2      'year-month-day
            typeToFormat = "yyyy/m/d"
        End Select
        If (sfType = "datetime") Then
            typeToFormat = typeToFormat + " h:mm" ' 5.15
        End If
        
    Case "string", "picklist", "phone" ' , "textarea"
      typeToFormat = "@"
    Case "currency"
      typeToFormat = "$#,##0_);($#,##0)" ' format as currency, no cents (added in 5.15)
    
    End Select
    
End Function

'
' check assumptions prior to an update
'
Function sfupdate_check(s As Range)
    sfupdate_check = False
    If (s.Areas.Count > 1) Then
      sfError "cannot run on multiple selections"
      Exit Function
      End If

    If (QueryRegBool(NOLIMITS)) Then
      sfupdate_check = True
      Exit Function
      End If
      
    ' adjust these limits to meet your requirements, or flip NOLIMITS in the options dialog
    If (s.Rows.Count > 3500 Or s.Columns.Count > 20) Then
      sfError "selection too large, cannot run on > 3500 rows and > 20 cols"
      Exit Function
    End If
    
    sfupdate_check = True
done:
End Function
'
' just a message no return value?
'
Function sfMissingTable(tbl As String)
    sfError "No table name found near " & _
    ActiveCell.Address & vbCrLf & "The " & tbl & _
    " command requires a valid Salesforce.com table name in the selected cell "
done:
End Function

'
' try to find a field in this object which is time based...
' used to format a query when none is found
'
Function recent_label(sfd As SObject4) As String
Dim c: For Each c In sfd.fields
    If c.name = "SystemModstamp" Then recent_label = c.label
Next c: If recent_label <> "" Then GoTo done
For Each c In sfd.fields
    If c.name = "LastModifiedDate" Then recent_label = c.label
Next c: If recent_label <> "" Then GoTo done
For Each c In sfd.fields
    If c.name = "CreatedDate" Then recent_label = c.label
Next c: If recent_label <> "" Then GoTo done
    
done:
End Function


'
' adjust the format of the value for types as expected by API
'
Function sfQueryValueFormat(typ, vlu)
   Select Case typ
      Case "datetime", "date":
        '
        ' 5.12 allow strings like
        '   today, today - 1 , today - 150, today + 30
        ' to be translated into vba dates for the query...
        '
        If (InStr(LCase(vlu), "today")) Then
          Dim today As Date: today = Date
          Dim daychange As Variant, incr%: incr = 0
          If (InStr(LCase(vlu), "-")) Then
            daychange = Split(vlu, "-")
            incr = 0 - Int(daychange(1))
          End If
          If (InStr(LCase(vlu), "+")) Then
            daychange = Split(vlu, "+")
            incr = Int(daychange(1))
          End If
          vlu = DateAdd("d", incr, today)
        End If ' 5.12 end
        
        sfQueryValueFormat = Format$(vlu, "yyyy-mm-ddTHH:MM:SS.000Z")
        
      Case "double", "currency", "percent":  ' add percent per Scot S. 5.67
        If (InStr(vlu, ".")) Then
            sfQueryValueFormat = Val(vlu) ' if the double has a decimal already, dont need to add .0
        Else
            sfQueryValueFormat = Val(vlu) & ".0"
        End If
      Case "boolean":
        sfQueryValueFormat = IIf((Val(vlu) Or "true" = LCase(vlu)), "TRUE", "FALSE")
      
      Case "int": ' 6.11 by scot stony
        sfQueryValueFormat = vlu
  
      Case Else: ' all which look like string, including but not limited to
        sfQueryValueFormat = "'" & vlu & "'" ' string, picklist, id, reference, textarea, combobox email
        
    End Select
End Function
'
' we may need to trim the range down to valid ID's ?
'
' trim top and bottom of this range to try to capture the region of valid
' object id's
' if we are given a range like "A:A" we can be smart by removing the
' top invalid items and triming the blank cells at the tail of the range
Function build_ref_range(str As String) As Range

    Dim t As Range, r As Range
    On Error Resume Next
    Set r = Range(str) ' if this is not a valid range description send a msg
    If r Is Nothing Then GoTo done ' Range method did not work
    On Error GoTo done
    
    Dim c As Range: For Each c In r
        If (c Is Nothing) Then GoTo done
        If IsError(c) Then GoTo done
        Select Case Len(c.value)
          Case 15, 18:
            ' sometimes a text string like 'opportunity id' will be just
            ' 15 or 18 long, to avoid adding this, check that the string we
            ' are looking at has some numeric chars and is not all alpha.
            If (c.value Like "*[0-9][0-9]*") Then    ' two adjacent numbers
                If (t Is Nothing) Then Set t = c ' special case first time thru
                Set t = Application.Union(t, c) ' normal case, extend the range down
                End If
        End Select
    Next c
    
    ' check that the range is made of one area...
     If t.Areas.Count > 1 Then
        MsgBox "Range " & t.Address & " is made of more than one area"
        End If
        
done:
    Set build_ref_range = t
End Function
'
' just writing an SObject4 into a row in excel, required globals are passed in
'
Sub format_write_row(g_body As Range, g_header As Range, _
  g_sfd As SObject4, so As SObject4, row As Integer)
  With g_body
  Dim maxRowHght: maxRowHght = ActiveSheet.StandardHeight * 3
  Dim j As Integer
    For j = 1 To g_header.Count
    Dim name$, fmt$, rheight%
    name = sfFieldLabelToName(g_sfd, g_header.Cells(1, j).value)
    fmt = typeToFormat(so.Item(name).Type)
    rheight = .Cells(row, j).RowHeight  ' before height
    
    ' map owner id to names (5.29)
    ' only do this if the option flag is set... or should it be default
    ' if querybool(SPELL_USERNAME) then ...
    '
    If so.Item(name).Type = "reference" Then
        .Cells(row, j).value = sfRefName(so.Item(name).value, so.Item(name).ReferenceTo)
    Else
    ' need to preserve text fields as text in excel or we may
    ' lose any leading zeros... !!!
    ' therefore we need to respect the field type here
    ' gotcha: the format must be set both before and after as the value
    ' assignment appears to trump some formats
    .Cells(row, j).NumberFormat = fmt
    ' 6.02 by MO'L
    ' Check type. Do not trim if it's a date or datetime as this will convert the date to text and lose international formatting: MO'L
    Select Case so.Item(name).Type
    Case "date", "datetime"
        .Cells(row, j).value = so.Item(name).value
    Case Else
        .Cells(row, j).value = Left(so.Item(name).value, 1023)
    End Select
    .Cells(row, j).NumberFormat = fmt  ' some formats like to be applied after

    If is_hyperlink(so.Item(name)) Then _
        Call add_hyperlink(.Cells(row, j), so.Item(name)) ' 6.09
    
    ' do something about the auto resizing, just to try to avoid blowup in long text
    ' fields as they are loaded into the cells, but dont mess with it if the user
    ' has set a height first
    If (.Cells(row, j).RowHeight > maxRowHght And _
      .Cells(row, j).RowHeight > rheight + 1) Then
      .Cells(row, j).RowHeight = maxRowHght ' set some default max
    End If
    
    End If
  Next j
  End With
End Sub


