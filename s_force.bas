Attribute VB_Name = "s_force"
Option Explicit
Dim salesforce As CSession ' holds a valid session

Dim g_sfd As SObject4      ' global describe for current table, not really needed since toolkit caches
Dim g_start As Range      ' some globals to hold info about the region we are working on
Dim g_ids As Range        ' column with salesforce ID's
Dim g_header As Range     ' row with column labels
Dim g_table As Range      ' all current region, with table name and header row
Dim g_body As Range       ' area with data, just below the header row
Dim g_objectType As String ' current entity table, ie "Account"
Dim g_labelNames As Scripting.Dictionary ' each table needs this mapped once to speed up drawing
' TODO make this batch size configurable via the options dialog box
Const maxBatchSize As Integer = 50  ' used to size Update,Query Row, Create and Delete batches
'// limits for update batch sizes
Public Const maxCols As Long = 20
Public Const maxRows As Long = 3500

'// enumerator for custom errors --> see sfUpdate_New() for details
Public Enum custErr
    err_doNothing = 5000
    err_noSession
    err_noSelection
    err_tooManyAreas
    err_tooManyRows
    err_tooManyCols
    err_noRowsf
    err_noCols
End Enum

'
' update --- update data from the worksheet into salesforce
'
Sub sfUpdate()
  If (QueryRegBool(SKIPHIDDEN)) Then
    Call sfUpdate_New ' 6.10
    Exit Sub
    End If
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done
  If (Selection.Rows.Count = 65536) Then ' entire row selected trim this
    Intersect(Selection, g_body).Select ' trim to region
  End If
  If Not message_user Then GoTo done
  If Not sfupdate_check(Selection) Then GoTo done ' some sanity checking
 
  ' batch up the data by walking the thru the selection range
  ' breaking it up into chunk size Ranges to be uploaded
  '
  Dim someFailed As Boolean: someFailed = False
  Dim row_pointer%: row_pointer = Selection.row ' row where we start to chunk
  Dim chunk As Range
  Do ' build a chunk Range which covers the cells we can update in a batch
    Set chunk = Intersect(Selection, ActiveSheet.Rows(row_pointer)) ' first row
    If chunk Is Nothing Then Exit Do          ' done here
    Set chunk = chunk.Resize(maxBatchSize) ' extend the chunk to cover our batchsize
    Set chunk = Intersect(Selection, chunk) ' trim the last chunk
    row_pointer = row_pointer + maxBatchSize ' up our row counter
   
    chunk.Interior.ColorIndex = 36
    
    If Not update_Range(chunk, someFailed) Then GoTo done ' do it
    
    Dim sav As Range: Set sav = Selection ' save the selection
    DoEvents: DoEvents: sav.Select ' and restore, allows a control-break in here
    
  Loop Until chunk Is Nothing
  
  If someFailed Then sfError "One or more of the selected rows could not be updated" & Chr(10) & _
        "see the comments in the colored cells for details"
  
  GoTo done
      
wild_error:
  MsgBox "Salesforce: Update() " & vbCrLf & _
        "invalid Range, missing data type, or other unknown error" & vbCrLf & _
        Error(), vbOKOnly + vbCritical
done:
  Application.StatusBar = False
End Sub

' for Update
' do the work on a given Range, grok the table layout, load globals
' call the update, indicate if any of these failed.
'
Function update_Range(todo As Range, someFailed As Boolean)
  update_Range = False
  Dim cSObject4 As SObject4 ' load the id, and the value to update
  Dim sd As Scripting.Dictionary
  Dim rec() As SObject4
  Dim idlist() As String
  Dim qr As QueryResultSet4
  Dim unInitVariant As Variant
  
  ReDim idlist(todo.Rows.Count - 1) ' counts are 1 based
  ReDim rec(UBound(idlist))
  
  Dim c, i%: i = 0
  For Each c In Intersect(g_ids, todo.EntireRow)
    idlist(i) = FixID(c.value): i = i + 1
  Next c
  If i = 0 Then GoTo done ' how ?
  
  Application.StatusBar = "Retrieve :" & _
    todo.row - Selection.row + 1 & " -> " & _
    todo.row - Selection.row + todo.Rows.Count & _
    " of " & CStr(Selection.Rows.Count)
  
  ' any fields which we will want to NULL out must be fetched first, oh my.
  Dim j, fields As String: fields = "Id"
  For i = 0 To UBound(idlist): For Each j In todo.Columns
    If todo.Offset(i, j.Column - todo.Column).Cells(1, 1).value = "" Then
        Dim fn As String: fn = sfFieldLabelToName(g_sfd, _
            g_header.Cells(1, 1 + j.Column - g_start.Column).value)
        If InStr(fields, fn) = 0 Then fields = fields & ", " & fn ' add it only once
    End If
  Next j: Next i
  
  ' ok, now we can fetch the id and fields required for the update
  Set qr = salesforce.Retrieve(fields, idlist, g_objectType)
  If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
  
  Set sd = salesforce.ProcessQueryResults(qr)
  
  For i = 0 To UBound(idlist)
    Set cSObject4 = sd.Item(idlist(i))
    Debug.Assert IsObject(cSObject4)
    
    For Each j In todo.Columns
      Dim lab As String
      lab = g_header.Cells(1, 1 + j.Column - g_start.Column).value
      Dim fld: fld = sfFieldLabelToName(g_sfd, lab)
      Dim target As Range
      Set target = todo.Offset(i, j.Column - todo.Column)
      cSObject4.Item(fld).value = toVBtype(target.Cells(1, 1), cSObject4.Item(fld))
      
      If (cSObject4.Item(fld).value = "") Then
        ' special case, must set value to an uninitialized variant
        ' to indicate to the COM toolkit that we want to "nil" the field
        ' toolkit translates this to passing "fieldsToNull" in the actual SOAP request
        cSObject4.Item(fld).value = unInitVariant
        
      End If
      
      ' deal with user names in a reference id field here 5.29
      ' and record types, and others that ref_id can deal with 5.34
      ' need to map a name into the actual ID prior to passing to update
      ' ref_id routine will return the passed in value if we don't map
      ' the ReferenceTo type provided (User,Group,Profile... etc) as a fallback
      If cSObject4.Item(fld).Type = "reference" Then
        cSObject4.Item(fld).value = salesforce.ref_id( _
            cSObject4.Item(fld).value, cSObject4.Item(fld).ReferenceTo)
      End If
      
    Next j
    cSObject4.Tag = target.row ' use the tag to store a row number on the worksheet
    Set rec(i) = cSObject4
  Next i
  
  Application.StatusBar = "Updating :" & _
    todo.row - Selection.row + 1 & " -> " & _
    todo.row - Selection.row + todo.Rows.Count & _
    " of " & CStr(Selection.Rows.Count)
    
  ' TODO set assignment rule on update for case and lead objects
  
  Call salesforce.DoUpdate(rec)
  '
  ' 5.43 go thru and mark the failed cells, send a message
  ' which will pop up at the end of all updates
  ' If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
  '
  Dim r As SObject4
  For i = 0 To UBound(rec)
    Set r = rec(i)
    Dim thisrow As Range, firstcel As Range
    Set thisrow = Intersect(todo.Offset(i, 1).EntireRow, todo)
    Set firstcel = thisrow.Offset(0, 0).Cells(1, 1)
    ' find out what is wrong with this record
    If (r.Error) Then
        'Debug.Print r.ErrorMessage
        ' turns out that if one field fails, the entire row fails
        thisrow.Interior.ColorIndex = 6
        For Each c In thisrow.Cells
            If Not (c.Comment Is Nothing) Then c.Comment.Delete
        Next c
        firstcel.AddComment
        firstcel.Comment.Text "Update Row Failed:" & Chr(10) & r.ErrorMessage
        firstcel.Comment.Shape.Height = 60 ' is this enough
        someFailed = True ' will message this later
    Else
        ' clear out the color on this row only
        ' also remove any comments which may now be incorrect
        ' for this entire row, need to clear on each col of the selection
        thisrow.Interior.ColorIndex = 0
        For Each c In thisrow.Cells
            If Not (c.Comment Is Nothing) Then c.Comment.Delete
        Next c
    End If
  Next i ' 5.43 end
  
  update_Range = True
done:
  Set sd = Nothing
  Set qr = Nothing
  
End Function
Function insert_Range(todo As Range, someFailed As Boolean)
  Dim createarray() As SObject4 ' holds array of objects to create in the range
  ReDim createarray(todo.Rows.Count - 1)
    
  Application.StatusBar = "Create :" & todo.row - Selection.row + 1 & " -> " & _
    todo.row - Selection.row + todo.Rows.Count & " of " & CStr(Selection.Rows.Count)
  Dim i: i = 0
  
  todo.Interior.ColorIndex = 36 ' show where we are working
  
  Dim rw: For Each rw In todo.Rows
    ' don't insert if there is no "new" label in the id column
    If Not (objectid(rw.row, True) Like "[nN][eE][wW]*") Then GoTo nextrw
 
    Set createarray(i) = salesforce.CreateEntity(g_objectType) ' creates an empty obj
    
    ' tag will specify which row in the sheet that we want to put the resulting ID back into
    ' allows us to skip rows which are not "new"...
    ' and allow us to walk the array later and place messages about which rows failed correctly
    createarray(i).Tag = rw.row
    
    Dim j, IDcol, tmpuid As String
    For j = 1 To g_header.Count
      Dim name As String
      name = sfFieldLabelToName(g_sfd, g_header.Cells(1, j).value)
      If name <> "Id" Then ' don't overwrite the id on this row, needs to be empty when passed to create
        ' find the field, TODO i have a routine to find the field, could refactor this loop.
        Dim fld
        For Each fld In createarray(i).fields
    
        If fld.name = name Then
            ' Application.StatusBar = "loading value for " & name
            If g_table.Cells(rw.row + 1 - g_table.row, j).value <> "" Then
              
              ' here we have a value check it and load it into the fld value
              ' 5.10 dont load field values unless the field is createable
              If (fld.Createable) Then
'                Dim vv
'                vv = g_table.Cells(rw.row + 1 - g_table.row, j).value
'                If (fld.Type = "currency") Then vv = vv & ".0"
                fld.value = toVBtype(g_table.Cells(rw.row + 1 - g_table.row, j), fld)
              End If
              
              ' 5.31 convert several reference id's into an name string if possible
              ' cleaned up in 5.43, 5.44
              If fld.Type = "reference" Then
                fld.value = salesforce.ref_id(fld.value, fld.ReferenceTo)
              End If
              
            End If
          End If
        Next fld
      Else
        IDcol = j ' save this location for later
      End If

    Next j
    i = i + 1 ' move to next array slot...
nextrw:
  Next rw

    ' resize createarray if it is shorter than we thought (todo.Rows.Count - 1)
    If i < 1 Then  ' no records to insert
        sfError "No records to Insert, enter the string 'New' on one or more rows"
        todo.Interior.ColorIndex = 0 ' clear out color
        GoTo done
    End If
    If i < todo.Rows.Count Then
        ReDim Preserve createarray(i - 1)
    End If
    
    ' case and lead may get a special header
    ' check options in reg first (5.17)
    Select Case LCase(g_objectType)
      Case "case"
        If QueryRegBool(CASERULE) Then Call salesforce.SetSoapHeader("AssignmentRuleHeader", "useDefaultRule", True)
      Case "lead"
        If QueryRegBool(LEADRULE) Then Call salesforce.SetSoapHeader("AssignmentRuleHeader", "useDefaultRule", True)
    End Select
    
    Call salesforce.DoCreate(createarray)
    todo.Interior.ColorIndex = 0 ' clear out color
    
    ' grab the id and place it into the cell that held the "new" specifier
    Dim r As SObject4
    For i = 0 To UBound(createarray): Set r = createarray(i)
        Dim cel As Range
        Set cel = g_table.Cells(r.Tag + 1 - g_table.row, IDcol)

        ' see if this row produced an error
        If r.Error Then ' put any error messages out in "comments" , clumsy but better than no message
            If Not (cel.Comment Is Nothing) Then cel.Comment.Delete
            cel.AddComment
            cel.Comment.Text "Insert Row Failed:" & Chr(10) & r.ErrorMessage
            cel.Comment.Shape.Height = 55
            cel.Interior.ColorIndex = 6 ' bright yellow
            someFailed = True ' will message this later
        Else ' good to go, put these ID's on the correct row !!!
            If Not (cel.Comment Is Nothing) Then cel.Comment.Delete
            cel.value = r.Item("Id").value
        End If
    Next i
    ' no return value from this func
done:
End Function
'
' load up a table given info in row 1 and 2 to construct a query
' finaly will use SOQL
'
Sub sfQuery(Optional RowsReturned As Integer) ' 6.12
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done ' gets a description
  
  With g_table
  
  ' 5.13 - allow special commands to be picked up from the normal query cells
  ' enables a clever join functionality using two adjcent tables
  ' so far the keywords that we pull out of this cell are :
  '   refresh
  '     if we see this keyword, rather than doing a query, we will just
  '     run sfQueryRows on all the existing rows,
  '     allows a left join to be built using two tables and an excel formula
  '   in
  '     use the field as a reference, use the value as a range
  '     query in a for loop across the range outer join?
  '   on
  '     like in, but pull only the first matching result so that this table
  '     will line up with the related table
  '
  Select Case LCase(Trim(g_table.Cells(1, 2).value))
    Case "refresh": Call sfRefresh: GoTo donedone ' did all we wanted by now
  End Select ' end 5.13
   
  ' if none, construct a reasonable query, like modified past 2 weeks...
  If (.Cells(1, 2).value = "") Then Call default_query(g_table)
  
  If (.Rows.Count > 2) Then ' remove old contents
    .Offset(2, 0).Resize(.Rows.Count - 2, .Columns.Count).Select
    Selection.ClearContents ' these are nice but leave old comments
    'Selection.ClearFormats ' done by the below
    'Selection.Clear ' clear out the old comments also  6.04
   
  End If
  g_start.Select
  
  Dim sels As String: sels = getSelectionList(g_sfd)
  
  Dim where$, jw%, lab$, opr$, field_obj As Object
  Dim vlu$, refIds As Range, outrow%, joinfield$
  Dim oneeachrow As Boolean: oneeachrow = False  ' for "on" joins
  jw = 2
  
  Do While (.Cells(1, jw).value <> "") ' if it's not empty, assume its more query
  
    lab = .Cells(1, jw).value ' the field label
    opr = .Cells(1, jw + 1).value ' the operator
    vlu = .Cells(1, jw + 2).value ' the criteria value(s)
    Set field_obj = sfField(g_sfd, lab) ' 5.46 get the field as an object
    
    ' operator
    ' add other aliases here if you like
    opr = LCase(opr): Select Case opr
      Case "equals": opr = "="
      Case "contains": opr = "like"
      Case "not equals": opr = "!="
      Case "less than": opr = "<"
      Case "less than or equals to": opr = "<="
      Case "greater than": opr = ">"
      Case "greater than or equals to": opr = ">="
    End Select
    
    ' 5.23 basic error check, 5.0 checks for this anyway but we know where the offending cell is
    If (opr = "like" And field_obj.Type = "picklist") Then
      sfError "like (or contains) operator in cell " & g_table.Cells(1, jw + 1).AddressLocal & _
      " is not valid on picklist fields, " & vbCrLf & " use --> equals, not equals"
      GoTo done
    End If
    
    ' special case 'in' and a ref field
    If ((opr = "in" Or opr = "on") And (field_obj.Type = "reference" Or field_obj.Type = "id")) Then
      If (opr = "on") Then oneeachrow = True
      Set refIds = build_ref_range(vlu) ' list of IDs to use in join
      joinfield = field_obj.name ' save for later, should be only one..
      
      If refIds Is Nothing Then
        sfError "Range error, could not build a valid range from the string" & _
          vbCrLf & "--> " & vlu & " <--" & vbCrLf & " in the cell " & .Cells(1, jw + 2).AddressLocal & _
          "expected valid range (ex: 'A:A') or range name"
        GoTo done
      End If
        
    Else ' general case
      ' Value ~ assemble the where clause using field, opr and values list
      '   this loop has been re-written (ver 5.04) to properly
      '   deal with comma seperated values i.e. -> field | operator |this,that|
      '   should become (field operator 'this' OR field operator 'that')
      '   unless it's multipicklist type then produce slightly different string for SOQL:
      '     (field inclqudes ('this') or field includes ('that'))
      '     (field excludes ('this') or field excludes ('that'))
      '
      Dim values As Variant: values = Split(vlu, ",")
      ' if values is empty and vlu is the nul string, still need to assemble the clause
      Dim vu, clause$: clause = "":
      If (UBound(values) < 0 And vlu = "") Then ' case of one empty value
          clause = field_obj.name & " " & opr & "''"
          If field_obj.Type = "date" Then ' special case, compare value is an empty date 5.49
            clause = field_obj.name & " " & opr & " null"
            End If
      Else
          For Each vu In values ' works for one or many non nul values
            
            Dim str$: str = vu:
            str = SFDC_escape_q(str) ' escape some chars
            vu = salesforce.ref_id(str, field_obj.ReferenceTo) ' map strs to refid's 5.46
            
            If Len(clause) > 0 Then
                If (opr Like "!=") Then
                    clause = clause & " and "
                Else
                    clause = clause & " or " ' prepend an or
                End If
            End If
            If (opr Like "like") Then vu = "%" & vu & "%"  ' wrap like string with wildcard
            If (opr Like "begins with" Or opr Like "starts with") Then vu = vu & "%" ' wildcar at front
            If (opr Like "ends with") Then vu = "%" & vu ' wlidcard at end
            If (opr Like "regexp") Then opr = "like"  ' pass the user provided wildcard
            Dim fmtVal$: fmtVal = sfQueryValueFormat(field_obj.Type, vu) ' format value for SOQL
            If (field_obj.Type = "multipicklist") Then fmtVal = "(" & fmtVal & ")" ' special case
            
            If (opr Like "starts with" Or opr Like "begins with" Or opr Like "ends with") Then  ' remap these to 'like'
                clause = clause & field_obj.name & " " & "Like" & " " & fmtVal '**** thanks to tim_bouscal!
            Else
                clause = clause & field_obj.name & " " & opr & " " & fmtVal ' assemble the clause
            End If
            
          Next vu
      End If
      
      If (UBound(values) > 0) Then clause = "(" & clause & ")" ' cant hurt
      where = where & clause ': Debug.Print where
      If (.Cells(1, jw + 3).value <> "") Then where = where & " and " ' to be ready for more
      
    End If
    
    jw = jw + 3 ' slide over to grab the next three cells
    
  Loop ' end loop while we have more WHERE clauses to add
  
  End With ' g_table
  
'  Debug.Print "select " & sels & " from " & g_objectType
'  Debug.Print " " & where
 
  ' to support join, if we saw a "reference in range" we need to loop over this,
  ' otherwise call once for a normal query
  outrow = 1 ' the row within g_body where output begins
  If refIds Is Nothing Then
    outrow = query_draw(sels, " where " & where, outrow, False) ' just one query
  Else
    Dim c As Range, tmp$, preJoinLen
    Dim in_values() As Variant
    ReDim in_values(0) As Variant
    
    ' 18 characters of control for query_draw ("select data from ", " ")
    ' 20 characters of control for where clause (" and ", " where ", " IN ('", "')")
    preJoinLen = Len(g_objectType) + Len(where) + Len(joinfield) + 18 + 20
    
    For Each c In refIds.Cells ' loop over a range to output a join
      in_values(UBound(in_values)) = c.value
      
      ' There is a limit of 10,000 characters in a query, let's try to make sure we
      ' aren't going to go over that by estimating the length of our query.
      ' 500 is pretty much the max as that's 9000 characters right there
      ' 2 to UBound(in_values) to see if the next id will set us over the limit
      ' 22 character per Id/Reference (18 characters from SFDC + "', '")
      ' ON queries require 1 select per row which is weak sauce
      ' TODO: Optimize ON queries
      If ((preJoinLen + ((UBound(in_values) + 2) * 21) >= 10000) Or _
        UBound(in_values) = 500) Or opr = "on" Then
        
        tmp = where
        If (where <> "" And Right(where, 4) <> "and ") Then tmp = where & " and " ' 5.56
        tmp = tmp & joinfield & " IN ('" & Join(in_values, "','") & "')" ' use the ID from the reference colum in each query
        outrow = query_draw(sels, " where " & tmp, outrow, oneeachrow)
        ReDim in_values(0) As Variant
      Else
        ReDim Preserve in_values(UBound(in_values) + 1) As Variant
      End If
    Next c
    
    ' Catch any that didn't fit previous queries
    If UBound(in_values) <> 1 And in_values(0) <> "" Then
      tmp = where
      If (where <> "" And Right(where, 4) <> "and ") Then tmp = where & " and " ' 5.56
      tmp = tmp & joinfield & " IN ('" & Join(in_values, "', '") & "')" ' use the ID from the reference colum in each query
      outrow = query_draw(sels, " where " & tmp, outrow, oneeachrow)
    End If
   
  End If
  If (outrow <= 1) Then sfError "No data returned for this Query"
  Set refIds = Nothing
  GoTo done
  
wild_error:
  sfError "Salesforce: Query() " & vbCrLf & _
    "invalid Range, missing data type, or other error, " & vbCrLf & _
    "type is: " & g_objectType & vbCrLf & Error()
done:
  Application.StatusBar = "Query : drawing complete, " & outrow - 1 & _
    " total rows returned"
donedone:
  Application.ScreenUpdating = True
  RowsReturned = outrow - 1 ' 6.12
  Application.StatusBar = False
End Sub
'
' aka: refresh
' select all valid rows in this table (top down) and call query rows on that
' can be used to build a Joined table in excel
'
Sub sfRefresh()
  On Error GoTo wild_error
  With g_body ' assumes we have already done session and set_ranges ok ?
  Dim rw: rw = 1
  ' special work to trim off rows where we know that the id is no good.
  ' assume these are at the bottom of the selection, start at top
  .Offset(0, 0).Resize(rw, .Columns.Count).Select ' select the first row
  Do
    If Not iSObject4Id(.Offset(rw).row) Then Exit Do  ' Offset is zero based Resize is not
    rw = rw + 1
    .Offset(0, 0).Resize(rw, .Columns.Count).Select ' select the next row
  Loop Until rw = .Rows.Count
  End With
  Call sfQueryRow
  g_start.Select ' nice touch
  GoTo done
wild_error:
  sfError "Salesforce: Query() " & vbCrLf & _
    "invalid Range, missing data type, or other error, " & vbCrLf & _
    "type is: " & g_objectType & vbCrLf & Error()
done:
End Sub

'
' describe a table on cells starting just under the table name, not from wizard
'
Sub sfDescribe()
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  On Error GoTo nosheet
  Set g_table = ActiveCell.CurrentRegion
  On Error GoTo wild_error
  Set g_start = g_table.Cells(1, 1)
  Set g_header = Range(g_table.Cells(2, 1), g_table.Cells(2, g_table.Columns.Count))
  g_objectType = g_start.value
  Dim fld
  With g_table
  If (.Rows.Count > 1) Then
    .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Select
    Selection.ClearContents ' read from the selection?
    g_start.Select
  End If
  
  Dim pos%
  If g_objectType = "" Then ' no object , just jump them into the quick start dialog
    ' however be sure to exit here as sfDescribe is called again
    ' from within the describeBox callback code
    If describeBox_init = True Then describeBox.Show
    GoTo done ' always
  End If
  
  ' Normal SObject4 describe case, not global
  Dim sobj As SObject4: Set sobj = salesforce.CreateEntity(g_objectType)
  If sobj Is Nothing Then
    sfMissingTable "could not find entity named >" & g_objectType
    If describeBox_init = True Then describeBox.Show ' launch the wizard if we didnt see a valid entity
    GoTo done
    End If
    
  pos = 0 ' since we did Id in the 0 pos
  For Each fld In sobj.fields ' drop the id in first col, sort of a tradition
    If fld.name = "Id" Then
        pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
    End If
  Next fld
  
  For Each fld In sobj.fields ' order the rest by placing required next
    If required(fld) Then
        pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
  Next fld
  
  For Each fld In sobj.fields ' named fields
    If namefield(fld) Then
    pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
    End If
  Next fld
  
  For Each fld In sobj.fields ' then standard
    If standard(fld) Then
        pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
  Next fld

  For Each fld In sobj.fields  ' now the custom fields, skipping id and std
    If custom(fld) Then
       pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
    End If
  Next fld
 
  For Each fld In sobj.fields ' read-only last ~ CreatedById,CreatedDate,LastModifiedById,...
    If readonly(fld) Then
       pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
    End If
  Next fld
  
  End With ' g_table
  
  GoTo done
wild_error:
    sfError "Salesforce: Describe() " & vbCrLf & _
    "invalid Range, missing data type, or other error, type is: " & _
    g_objectType & vbCrLf & Error()
    GoTo done
    
nosheet:
  MsgBox "Oops, Could not find an active Worksheet"
  
done:
End Sub
'
' describe a table load into select list, called from wizard
' called before step 2b to describe the fields and load into the picklist
' added for 5.63
'
Sub sfDescribe_wiz()
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  On Error GoTo nosheet
  Set g_table = ActiveCell.CurrentRegion
  On Error GoTo wild_error
  Set g_start = g_table.Cells(1, 1)
  Set g_header = Range(g_table.Cells(2, 1), g_table.Cells(2, g_table.Columns.Count))
  g_objectType = g_start.value
  
  With g_table
  If (.Rows.Count > 1) Then
    .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Select
    Selection.ClearContents ' read from the selection?
    g_start.Select
  End If
  
'  If g_objectType = "" Then ' no object , just jump them into the quick start dialog
'    ' however be sure to exit here as sfDescribe is called again
'    ' from within the describeBox callback code
'    If describeBox_init = True Then describeBox.Show
'    GoTo done ' always
'  End If
  
  ' Normal SObject4 describe case, not global
  'Dim sobj As SObject4: Set sobj = salesforce.CreateEntity(g_objectType)
  Set g_sfd = salesforce.CreateEntity(g_objectType)
  If g_sfd Is Nothing Then
    sfMissingTable "could not find entity named >" & g_objectType
    If describeBox_init = True Then describeBox.Show ' launch the wizard if we didnt see a valid entity
    GoTo done
  End If
    
  
  Dim fld
  Step2b.Step2bField.Clear ' clean it out first
  '
  ' add these in the same order we would draw them on the row
  '
  For Each fld In g_sfd.fields ' drop the id in first col, sort of a tradition
    If fld.name = "Id" Then Step2b.Step2bField.AddItem fld.label
  Next fld
  For Each fld In g_sfd.fields ' order the rest by placing required next
    If required(fld) Then Step2b.Step2bField.AddItem fld.label
  Next fld
  For Each fld In g_sfd.fields ' named fields
    If namefield(fld) Then Step2b.Step2bField.AddItem fld.label
  Next fld
  For Each fld In g_sfd.fields ' then standard
    If standard(fld) Then Step2b.Step2bField.AddItem fld.label
  Next fld
  For Each fld In g_sfd.fields  ' now the custom fields, skipping id and std
    If custom(fld) Then Step2b.Step2bField.AddItem fld.label
  Next fld
  For Each fld In g_sfd.fields ' read-only last ~ CreatedById,CreatedDate,LastModifiedById,...
    If readonly(fld) Then Step2b.Step2bField.AddItem fld.label
  Next fld
  
  End With
  GoTo done
wild_error:
    sfError "Salesforce: Describe() " & vbCrLf & _
    "invalid Range, missing data type, or other error, type is: " & _
    g_objectType & vbCrLf & Error()
    GoTo done
  
nosheet:
done:
End Sub
'
' just draw the labels from the selected list of the wizard, then return
' really part of the wizard, only called from step 2b and before step 3
' show and hides are done in the dialog code blocks to make it easy to debug.
' all we do here is draw the labels that were selected in the listbox
' this fuction added for 5.63
'
Sub sfDescribe_wiz_draw()
  Dim toscreen As Boolean
  Dim j As Integer
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  On Error GoTo nosheet
  Set g_table = ActiveCell.CurrentRegion
  On Error GoTo wild_error
  Set g_start = g_table.Cells(1, 1)
  Set g_header = Range(g_table.Cells(2, 1), g_table.Cells(2, g_table.Columns.Count))
  g_objectType = g_start.value
  Dim fld
  Dim pos%
  With g_table
  Dim sobj As SObject4: Set sobj = salesforce.CreateEntity(g_objectType)
  For Each fld In sobj.fields ' drop the id in first col, sort of a tradition
    If fld.name = "Id" Then
        If (Step2b.is_wiz_selected(fld.label)) Then
                pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld
 
  For Each fld In sobj.fields ' order the rest by placing required next
    If required(fld) Then
        If (Step2b.is_wiz_selected(fld.label)) Then
            pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld
  
  For Each fld In sobj.fields ' named fields
    If namefield(fld) Then
        If (Step2b.is_wiz_selected(fld.label)) Then
                pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld
  
  For Each fld In sobj.fields ' then standard
    If standard(fld) Then
        If (Step2b.is_wiz_selected(fld.label)) Then
            pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld

  For Each fld In sobj.fields  ' now the custom fields, skipping id and std
    If custom(fld) Then
         If (Step2b.is_wiz_selected(fld.label)) Then
            pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld
 
  For Each fld In sobj.fields ' read-only last ~ CreatedById,CreatedDate,LastModifiedById,...
    If readonly(fld) Then
         If (Step2b.is_wiz_selected(fld.label)) Then
            pos = decorateColumnHeader(g_start.Offset(1, pos), fld, pos)
        End If
    End If
  Next fld
  
  End With ' g_table
  
  GoTo done
wild_error:
    sfError "Salesforce: Describe() " & vbCrLf & _
    "invalid Range, missing data type, or other error, type is: " & _
    g_objectType & vbCrLf & Error()
    GoTo done
    
nosheet:
  MsgBox "Oops, Could not find an active Worksheet"
  
done:
End Sub
' what do we know, well these appear to be required
Private Function required(fld)
  required = fld.name <> "Id" And fld.Nillable = False And fld.DefaultOnCreate = False And fld.Createable = True
End Function
Private Function namefield(fld)
  namefield = fld.namefield And Not required(fld) And Not (fld.custom) And fld.Updateable ' Id is not updateable
End Function
Private Function standard(fld)
  standard = Not namefield(fld) And Not required(fld) And Not (fld.custom) And fld.Updateable ' Id is not updateable
End Function
Private Function custom(fld)
  ' 16.03 exclude required, those are already on the screen, even if the are custom...
  custom = fld.custom And fld.Updateable And Not required(fld)
End Function
Private Function readonly(fld)
  readonly = fld.name <> "Id" And Not (fld.Updateable) And Not required(fld)
End Function
'
' 5.04 ~ given a cell and a field, label and decorate it
' also increment the position
Private Function decorateColumnHeader(cel As Range, fld, pos)
  cel.value = fld.label ' first thing first

  If (QueryRegBool(USE_NAMES)) Then cel.value = fld.name ' 5.24
  cel.WrapText = True
  decorateColumnHeader = pos + 1
  
  ' clear it out and left over comments
  If Not (cel.Comment Is Nothing) Then cel.Comment.Delete
  Dim commentStr$, commentHeight%: commentHeight = 60 ' default holds ~6 lines
  ' do some fancy formating with comment fields

  If Not fld.Updateable Then commentStr = "Read Only Field" & Chr(10)
  If required(fld) Then commentStr = "Required on Insert" & Chr(10)
  If fld.name = "Id" Then commentStr = "Primary Object Identifier" & Chr(10)
  ' 5.57 add name
  commentStr = commentStr & "API Name: " & fld.name & Chr(10)

  Select Case fld.Type
   Case "picklist", "multipicklist":  ' OR multipicklist ???
    commentStr = commentStr & "Type: " & fld.Type & Chr(10) & _
      (Join(fld.PickListValues, Chr(10)))
    Dim h%: h = (UBound(fld.PickListValues) * 12) + 1
    If (h > 60) Then commentHeight = h
    Case Else
    commentStr = commentStr & "Type: " & fld.Type & Chr(10)
  End Select
  ' other decorations are possible
  ' dim out read-only cols?
  ' color red for required fields, or bold?
  If commentStr = "" Then Exit Function
  cel.AddComment
  cel.Comment.Text commentStr
  cel.Comment.Shape.Height = commentHeight
End Function

Sub sfInsertRow()
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done
  
  If Not QueryRegBool(GOAHEAD) Then
    Dim msg ' warn the user that they are writing to the database
    msg = "You are about to INSESRT: " & CStr(Selection.Rows.Count) & " " & g_objectType & _
     " record(s) in to Salesforce.com"
    msg = msg & vbCrLf
    If (MsgBox(msg, vbApplicationModal + vbOKCancel + vbExclamation + vbDefaultButton1, _
       "-- Ready to Update Salesforce.com --") = vbCancel) _
       Then GoTo done
  End If
  
  ' trim selection if it includes blanks...?
  ' 5.36
  Dim row_pointer%: row_pointer = Selection.row ' row where we start to chunk
  Dim chunk As Range, someFailed As Boolean: someFailed = False
  Do ' build a chunk Range which covers the cells we can create in a batch
    Set chunk = Intersect(Selection, ActiveSheet.Rows(row_pointer)) ' first row
    If chunk Is Nothing Then Exit Do          ' done here
    Set chunk = chunk.Resize(maxBatchSize) ' extend the chunk to cover our batchsize
    Set chunk = Intersect(Selection, chunk) ' trim the last chunk
    row_pointer = row_pointer + maxBatchSize ' up our row counter
'    chunk.Interior.ColorIndex = 36
    
    Call insert_Range(chunk, someFailed) ' doit
    
 '   chunk.Interior.ColorIndex = 0
    Dim sav As Range: Set sav = Selection ' save the selection
    DoEvents: DoEvents: sav.Select ' and restore, allows a control-break in here
    
  Loop Until chunk Is Nothing
  
  If someFailed Then sfError "One or more of the selected rows could not be inserted" & Chr(10) & _
        "see the comments in the colored cells for details"
        
  GoTo done
wild_error:
   sfError "Salesforce: Create() " & vbCrLf & _
    "invalid Range, missing data type, or other error, " & vbCrLf & _
    "type is: " & g_objectType & vbCrLf & Error()
done:
' comment on any that may have failed?? TODO
Application.StatusBar = "Insert Selected Rows completed"
End Sub


'
' look at each id object column, read column names from above
' construct a query and return that row from the database
' repeat on any selected rows
'
' needs to work in a batch mode, only query maxbatchsize each time
' retreive appears to hang sometimes, so we keep the batch size small
' and allow a "save file" to be pushed while we are working...
'
Sub sfQueryRow()
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done

  Dim sels As String, i As Integer
  sels = getSelectionList(g_sfd)
  
  If (Selection.Rows.Count = 65536) Then ' entire row selected trim this
    Intersect(Selection, g_body).Select ' trim to region
  End If
  
  
  If Not QueryRegBool(NOLIMITS) Then
    If (Selection.Rows.Count > 1999) Then
      sfError "too many rows selected " & Selection.Rows.Count & "max is 2000"
      GoTo done
    End If
  End If
  i = 0
  
  Dim row_pointer%: row_pointer = Selection.row ' row where we start to chunk
  Dim chunk As Range
  Do ' build a chunk Range which covers the cells we can query in a batch
    Set chunk = Intersect(Selection, ActiveSheet.Rows(row_pointer))
    If chunk Is Nothing Then Exit Do
    Set chunk = chunk.Resize(maxBatchSize) ' extend the chunk to cover our batchsize
    Set chunk = Intersect(Selection, chunk) ' trim the last chunk !
    row_pointer = row_pointer + maxBatchSize ' up our pos counter
'    Debug.Print "here do the chunk, " & chunk.row & " " & chunk.Rows.Count
'    For Each c In chunk
'      Debug.Print c.value
'    Next c
 
    chunk.Interior.ColorIndex = 36 ' show off...
    Application.ScreenUpdating = False
    If Not query_row(sels, chunk) Then GoTo done ' do it
    Application.ScreenUpdating = True
    chunk.Interior.ColorIndex = 0
    
  Loop Until chunk Is Nothing
    
  GoTo done
wild_error:
      sfError "Salesforce: Update() " & vbCrLf & _
        "invalid Range, missing data type, or other error, type is: " & _
        g_objectType & vbCrLf & Error()
done:
Application.StatusBar = False
End Sub
' do the query and fill in a range
' called from sfQueryRow
'
Function query_row(sels As String, todo As Range)
  query_row = False
  Dim i%: i = 0
  Dim idlist() As String: ReDim idlist(todo.Rows.Count - 1)
  Dim rw: For Each rw In todo.Rows
    idlist(i) = objectid(rw.row)
    i = i + 1
  Next rw
  
  Application.StatusBar = "Query records from salesforce>" & UBound(idlist) + 1
  Dim qr As QueryResultSet4: Set qr = salesforce.Retrieve(sels, idlist, g_objectType)
  If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done

  Dim sd As Scripting.Dictionary
  Set sd = salesforce.ProcessQueryResults(qr) ' just make restults into dict
  Application.StatusBar = "back from retrieve data at salesforce"
  i = 1
  For Each rw In todo.Rows
    todo.Rows(i).Interior.ColorIndex = 36
    Dim so As SObject4
    Set so = sd.Item(objectid(rw.row))
    Debug.Assert IsObject(so)
    Dim j: For j = 1 To g_header.Count
      Dim name As String
      name = sfFieldLabelToName(g_sfd, g_header.Cells(1, j).value)
      If name <> "Id" Then ' don't overwrite the id on this row.. it may be a formula
        ' write some fields after a lookup 5.29
        
        ' TODO this should be dropped into a function
        ' as we are duplicating code here and in format_write_row
        If (so.Item(name).Type = "reference" And _
            sfRefName(so.Item(name).value, so.Item(name).ReferenceTo) <> NOT_FOUND) Then
            g_table.Cells(rw.row + 1 - g_table.row, j).value = _
                sfRefName(so.Item(name).value, so.Item(name).ReferenceTo) ' given id, return name
        Else
            ' resepct the data type
            ' also use NumberFormatLocal if set, as this will preserve
            ' international currency
            Dim fmt As String
            fmt = typeToFormat(so.Item(name).Type)
            g_table.Cells(rw.row + 1 - g_table.row, j).NumberFormat = fmt
            ' avoid mem crash
            g_table.Cells(rw.row + 1 - g_table.row, j).value = Left(so.Item(name).value, 1023)
            g_table.Cells(rw.row + 1 - g_table.row, j).NumberFormat = fmt
            
            If is_hyperlink(so.Item(name)) Then _
                Call add_hyperlink(g_table.Cells(rw.row + 1 - g_table.row, j), so.Item(name)) ' 6.09

        End If
        
      End If
    Next j
    todo.Rows(i).Interior.ColorIndex = 0
    i = i + 1
  Next rw
  query_row = True
done:
  Set sd = Nothing
  Set qr = Nothing
End Function
'
'
' not tested
'
Sub sfDelete()
  Dim chunk As Range
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done
    
  ' entire row selected trim this to region
  If (Selection.Rows.Count = 65536) Then Intersect(Selection, g_body).Select
    
  Dim msg$: msg = "You are about to DELETE: " & CStr(Selection.Rows.Count) & " " & g_objectType & _
   " record(s)" & vbCrLf
  If (MsgBox(msg, vbApplicationModal + vbOKCancel + vbExclamation + vbDefaultButton1, _
      "-- Ready to Update Salesforce.com --") = vbCancel) Then GoTo done
  
  Dim row_pointer%: row_pointer = Selection.row ' row where we start to chunk
  Do ' build a chunk Range which covers the cells we can update in a batch
    Set chunk = Intersect(Selection, ActiveSheet.Rows(row_pointer))
    If chunk Is Nothing Then Exit Do
    
    Set chunk = chunk.Resize(maxBatchSize) ' extend the chunk to cover our batchsize
    Set chunk = Intersect(Selection, chunk) ' trim the last chunk !
    row_pointer = row_pointer + maxBatchSize ' up our pos counter
    
'    Debug.Print "here do the chunk, " & chunk.row & " " & chunk.Rows.Count
'    For Each c In chunk
'      Debug.Print c.value
'    Next c
    chunk.Interior.ColorIndex = 36 ' show off...
    If Not delete_Range(chunk) Then GoTo done ' do it
    chunk.Interior.ColorIndex = 0
    
  Loop Until chunk Is Nothing
  
  GoTo done
wild_error:
  sfError "Salesforce: Delete() " & vbCrLf & _
        "invalid Range, missing data type, or other error, type is: " & _
        g_objectType & vbCrLf & Error()
done:
  Application.StatusBar = False
End Sub

Function delete_Range(todo As Range)
  Dim idlist() As String
  ReDim idlist(todo.Rows.Count - 1)
  Dim i, rw: For Each rw In todo.Rows
    idlist(i) = objectid(rw.row)
    i = i + 1
  Next rw

  Application.StatusBar = "Delete :" & _
    todo.row - Selection.row + 1 & " -> " & _
    todo.row - Selection.row + todo.Rows.Count & _
    " of " & CStr(Selection.Rows.Count)
  
  Call salesforce.DoDelete(idlist, g_objectType)
  If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
  
  ' TODO when one fails we don't need to skip all of the rest...
  ' do it like the create calls?
  
  For Each rw In todo.Rows ' draw results
      Intersect(g_ids, ActiveSheet.Rows(rw.row)).value = "deleted"
  Next rw
  delete_Range = True
done:
  
End Function
'
' verify the session or login and create one
' clean up in 5.28, simplify...
'
Public Function SessionMe()
  SessionMe = False
  Set salesforce = Login.ensureSalesforce() ' get server value from login module
  If (salesforce Is Nothing) Then GoTo done
  SessionMe = True
done:
End Function

Private Function recent_label(sfd As SObject4) As String
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
Private Function getSelectionList(sfd As SObject4) As String
Dim c, sels As String
For Each c In g_header.Cells
  ' Debug.Print "cell header is >" & c.value
    If c.value <> "" Then
      Dim name As String: name = sfFieldLabelToName(sfd, c.value)
      If name <> "" Then sels = sels & name & ", "
    End If
  Next c
  sels = RTrim(sels)
  If (Mid(sels, Len(sels), 1) = ",") Then sels = Left(sels, Len(sels) - 1)  ' remove the final comma
  getSelectionList = sels
End Function
Function sfNameToVal(sfd As SObject4, name As String)
  Dim c: For Each c In sfd.fields
    If c.name = name Then
      sfNameToVal = c.value
      Exit Function
    End If
  Next c
End Function

' given an id, return a string of the user's name
' used for owner lookup,
' calls routines in csession which cache queries
Function sfUserName(userid As String) As String
    Call SessionMe
    sfUserName = salesforce.ref_name(userid, "User")
End Function
' query the user id each time given the full name
Function sfUserId(fullname As String) As String
    Call SessionMe: sfUserId = salesforce.ref_id(fullname, "User")
End Function
Function sfRefName(userid As String, t As String) As String
    Call SessionMe: sfRefName = salesforce.ref_name(userid, t)
End Function
Function sfRecTypeName(id As String) As String
 Call SessionMe: sfRecTypeName = salesforce.ref_name(id, "RecordType")
End Function
Function sfRecTypeId(na As String) As String
 Call SessionMe: sfRecTypeId = salesforce.ref_id(na, "RecordType")
End Function
Function sfAccountField(lic As String, field As String) As String
    Dim lic_d As Scripting.Dictionary
    sfAccountField = "#N/A"
    If Not SessionMe Then GoTo done
    On Error Resume Next
    If Not (lic_d.Count > 1) Then
      Dim queryData As QueryResultSet4
      
      Set queryData = salesforce.query( _
      "select Id from Account where " & field & " = " & "'" & lic & "'")
        
      If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
      Set lic_d = salesforce.ProcessQueryResults(queryData)
    End If
    
    Dim k: For Each k In lic_d.Keys
      sfAccountField = sfNameToVal(lic_d(k), "Id")
    Next k
done:
End Function

Function sfLogout()
    Set salesforce = Nothing: Call Login.logout
End Function

Function sfFieldType(sfd As SObject4, label As String) As String
    sfFieldType = "string" 'default
    Dim f As Integer
    For f = LBound(sfd.fields) To UBound(sfd.fields)
      Dim lc: lc = LCase(label)
        If (lc = LCase(sfd.fields(f).label) Or lc = LCase(sfd.fields(f).name)) Then
            sfFieldType = sfd.fields(f).Type
            Exit For
        End If
    Next f
'  5.05  If (sfFieldType = "") Then
'        sfError "invalid label -->" & label
'    End If
End Function
' given a label, return the field definition 5.46
Function sfField(sfd As SObject4, label As String) As Object
    Dim f As Integer
    For f = LBound(sfd.fields) To UBound(sfd.fields)
      Dim lc: lc = LCase(label)
        If (lc = LCase(sfd.fields(f).label) Or lc = LCase(sfd.fields(f).name)) Then
            Set sfField = sfd.fields(f)
            Exit For
        End If
    Next f
    
End Function
'
' was too slow to call in a drawing loop
' must map the labels to names somewhere and just grab the name... like a dict...
' this now builds a global labelNames dict, should be much faster...
'
Function sfFieldLabelToName(sfd As SObject4, label As String) As String
  sfFieldLabelToName = g_labelNames.Item(label)
  If sfFieldLabelToName = "" Then
    sfFieldLabelToName = sfFieldLabelToName_slow(sfd, label) ' look it up and add to dict
    If (Not (g_labelNames.Exists(label))) Then
        g_labelNames.Add label, sfFieldLabelToName  ' stash it
    End If
    ' only if it's not there
  End If
End Function
' called once on the header row to build the dict
Function sfFieldLabelToName_slow(sfd As SObject4, label As String) As String
' given the label, return the name
    sfFieldLabelToName_slow = ""
    label = LCase(label)
    Dim lent As String: lent = LCase(g_objectType)
    Dim f As Integer
    For f = LBound(sfd.fields) To UBound(sfd.fields)
      Dim lc As String
      lc = LCase(sfd.fields(f).label)
   '   Debug.Print "label is " & label & " label is " & lc & ", name is " & sfd.fields(f).name
      If (lc = label Or lc = CStr(lent & " " & label) Or _
        label = LCase(sfd.fields(f).name)) Then   ' do a lower case compare, allow raw names to match
         sfFieldLabelToName_slow = sfd.fields(f).name
         GoTo done
      End If
      
      ' match recordType id to something
      ' remove spaces and try again...
      Dim tmp As String: tmp = Replace(label, " ", "")
      If (tmp = CStr(lent & "id")) Then
      '  Debug.Print "label is " & label & " label is " & lc & ", name is " & sfd.fields(f).name
        sfFieldLabelToName_slow = "Id"
        GoTo done
      End If
      
    Next f
    
    If (sfFieldLabelToName_slow = "") Then
      ' allow for a special prefix handy for calc fields...
      If Left(label, 2) <> "x " Then
        ' Debug.Assert False
        sfError "invalid label -->" & label
      End If
    End If
done:
   ' Debug.Print "sfFieldLabelToName_slow label was " & label & ", return is " & sfFieldLabelToName
End Function
'
' called once early on, to figure out which column has the real id's
' column number is relative to g_header
'
Function getObjectIdColumn(sfd As SObject4) ' have a map of labels and one is the id
  Dim f, idLabel As String
  For Each f In sfd.fields
    If f.name = "Id" Then idLabel = LCase(f.label)
  Next f
  Debug.Assert idLabel <> "" ' better have an id or something is wrong
  
  Dim j As Integer: For j = 1 To g_header.Count
    'Debug.Print g_header.Cells(1, j).value
    With g_header.Cells(1, j)
    If idLabel = LCase(.value) Then
      getObjectIdColumn = j
      Exit Function
    End If
    ' custom objects exported using the excel report tool from salesforce
    ' have a form like this "label: ID" --> "Edi Set-up: ID"
    ' not the expected "record id", so we need to look for ": id" also
    If (Right(.value, 4) = ": ID" Or .value = "Id") Then
      getObjectIdColumn = j
      Exit Function
    End If
    End With
  Next j
  
  MsgBox "no Object Id found in the column header row"
End Function

'
' set global ranges, labels for the current active region
' used by most other calls which operate on a Range of data
' except sfDescribe which creates a default table layout
'
Function set_Ranges()
  Application.StatusBar = "build data Ranges"
  set_Ranges = False
  On Error GoTo nosheet
  Set g_table = ActiveCell.CurrentRegion
  On Error GoTo wild_error
  Set g_start = g_table.Cells(1, 1)
  ' see how many rows we have before setting body... 5.20
  If g_table.Rows.Count = 2 Then
    ' body is going to be outside the table... so place it where we need
    Set g_body = Range(g_table.Cells(3, 1).AddressLocal) ' 5.20
  Else
    Set g_body = Range(g_table.Cells(3, 1), g_table.Cells(g_table.Rows.Count, g_table.Columns.Count))
  End If
  
  'Debug.Print "g_table is at " & g_table.AddressLocal
  'Debug.Print "g_body is at " & g_body.AddressLocal
  
  ' trim the g_header Range down if g_table.columns.count is greater than
  ' the number of non blank cells in row 2 !!
  Dim k
  For k = 1 To g_table.Columns.Count
    If g_table.Cells(2, k) <> "" Then ' expand the Range to hold non blank cells
      Set g_header = Range(g_table.Cells(2, 1), g_table.Cells(2, k))
    End If
  Next k
  
  g_objectType = g_start.value
  If (g_objectType = "") Then
    'If describeBox_init = True Then describeBox.Show
    sfError "could not locate a object name in cell " & g_start.Address & vbCrLf & _
     "use Describe Sforce Object menu item to select a valid object"
    set_Ranges = False
    Exit Function
  End If
  
  Application.StatusBar = "Query " & g_objectType & " table description"
  Set g_sfd = salesforce.CreateEntity(g_objectType) ' requires we have logged in
  ' if g_sfd is nothing perhaps we need to figure out what type
  ' of table we are looking at,
  If (g_sfd Is Nothing) Then
    g_objectType = guess_table(g_table) ' assume top row is columns
    If (g_objectType = "") Then
        sfMissingTable "could not guess table in Query, Update, Delete or Insert Row"
        GoTo done
    Else ' found a table name in the column header fix up the globals...
        ' g_body
        Set g_body = Range(g_table.Cells(2, 1), _
                        g_table.Cells(g_table.Rows.Count, g_table.Columns.Count))
        ' g_objectType is already set
        Set g_header = Range(g_table.Cells(1, 1), g_table.Cells(1, g_table.Columns.Count))
        Set g_sfd = salesforce.CreateEntity(g_objectType)
        ' then we can continue as if we knew what was going on....
    End If
  End If
  Dim gcol
  gcol = getObjectIdColumn(g_sfd)
  Set g_ids = Intersect(g_body, g_body.Columns(gcol))
    
  ' optimization for looking up names from labels
  If g_labelNames Is Nothing Then ' init the dict
    Set g_labelNames = CreateObject("Scripting.Dictionary")
  Else  ' clean it up first
    g_labelNames.RemoveAll
  End If
  ' now fill it
  Dim c: For Each c In g_header
    If Not g_labelNames.Exists(c.value) Then
        Dim name As String: name = sfFieldLabelToName_slow(g_sfd, c.value)
        If name <> "" Then ' dont add blanks
          g_labelNames.Add c.value, name
        End If
    End If
  Next c
  
  set_Ranges = True
  
  GoTo done
wild_error:
  sfError "Salesforce: Query() " & vbCrLf & _
    "invalid Range, missing data type, or other error, " & vbCrLf & _
    "type is: " & g_objectType & vbCrLf & Error()
  GoTo done
nosheet:
  MsgBox "Oops, Could not find an active Worksheet"
done:
'  Debug.Print "gbody is at " & g_body.AddressLocal
End Function

' warn the user that they are writing (update) to the database
Public Function message_user()

  If (QueryRegBool(GOAHEAD)) Then
    message_user = True
    Exit Function
    End If
  
  Dim msg, cl, rw
  msg = "You are about to update: " & CStr(Selection.Rows.Count) & " " & _
  g_objectType & _
   " record(s)" & vbCrLf & " ---  Record ID      --- "
  For Each cl In Selection.Columns
   msg = msg & g_header.Cells(1, 1 + cl.Column - g_start.Column).value & ", "
  Next cl
  msg = msg & vbCrLf
    
  For Each rw In Selection.Rows
      msg = msg & objectid(rw.row)
      For Each cl In Selection.Columns
          msg = msg & ", " & CStr(g_start.Offset(rw.row - g_start.row, _
              cl.Column - g_start.Column).value)
      Next cl
      msg = msg & vbCrLf
      ' check the length of msg here, avoid a bomb out
      If (Len(msg) > 512) Then GoTo showmsg
  Next rw

showmsg: ' would be nice in fixed font
  If (MsgBox(msg, vbApplicationModal + vbOKCancel + vbExclamation + vbDefaultButton1, _
      "-- Ready to Update Salesforce.com --") = vbCancel) Then
      message_user = False
      Else
      message_user = True
      End If
done:
End Function

Function label_value(so As SObject4, lab As String)
  Dim name As String
  name = sfFieldLabelToName(g_sfd, lab)
  label_value = Left(so.Item(name).value, 1023)  ' avoid mem crash
End Function

'
' do a search of an string in sf.com, return the id
' warning: findfirst will give something back, just not always the one you want.
Function sfSearch(table, cell_or_string, Optional which_fields, Optional findfirst)
  sfSearch = "#N/F"
  Application.Volatile (False)
  
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  Dim tmp As String
  ' see if cel is a range or string..
  If VarType(cell_or_string) = vbString Then
    tmp = cell_or_string
  Else ' its a range... we hope
    tmp = RTrim(cell_or_string.Offset(0, 0).value)
  End If
  DoEvents
  If IsMissing(which_fields) Then which_fields = "NAME" ' 5.58
  ' debug.print "find {" & SFDC_escape(tmp) & "} in " & _
  ' which_fields & " fields RETURNING " & table & "(id) "
    
  Dim d As Scripting.Dictionary
  Set d = salesforce.Search("find {" & SFDC_escape(tmp) & "} in " & _
    which_fields & " fields RETURNING " & table & "(id) ")
  ' now which_fields must be "name" or "email" or a valid field in the table
'
  If Not salesforce.Msg_and_Fail_if_APIError Then GoTo done
  If d Is Nothing Then Exit Function ' no data found

  DoEvents ' nice, but is it needed ?

  ' if the dict found contains exactly one, return the ID
  Select Case d.Count
    Case 1
      sfSearch = d.Keys(0)
    Case Is > 1
      If Not IsMissing(findfirst) Then
        sfSearch = d.Keys(0)
      Else
        sfSearch = "found " & d.Count ' else the id is misleading, so dont guess
        ' could put out an array ?
      End If
  End Select
 
  
GoTo done
wild_error:
      sfError "Salesforce: Search() " & vbCrLf & _
        "invalid Range, missing data type, or other error: " & _
         vbCrLf & Error()
done:
End Function

'
' do a select using a passed in query string, return the id
'
Function soql_table(table, cel)
  Application.Volatile (False)
  On Error GoTo wild_error
  If Not SessionMe Then GoTo done
  Dim tmp As String
  ' see if cel is a range or string..
  If VarType(cel) = vbString Then
    tmp = cel
  Else ' its a range... we hope
    tmp = RTrim(cel.Offset(0, 0).value)
  End If

  soql_table = "#N/F"
  
  Dim queryData As QueryResultSet4
  Dim qry As String: qry = "select id from " & table & " where "
  Set queryData = salesforce.query(qry & tmp)
  
  If Not salesforce.Msg_and_Fail_if_APIError Then GoTo done
  Dim d As Scripting.Dictionary
  Set d = salesforce.ProcessQueryResults(queryData)
  
  If d Is Nothing Then Exit Function
    
  ' if the dict found contains exactly one, return the ID
  Select Case d.Count
    Case 1
      soql_table = d.Keys(0) ' only one key, return it
    Case Is > 1
      ' TODO could build an array incase we are looking for multiple
      soql_table = "found " & d.Count ' else the id is misleading, so dont guess
  End Select
  
GoTo done
wild_error:
      sfError "Salesforce: soql query() " & vbCrLf & _
        "invalid Range, missing data type, or other error: " & _
         vbCrLf & Error() & vbCrLf & qry
done:
End Function
'
' part of our first wizard
'
Public Function describeBox_init()
 On Error GoTo wild_error
 describeBox_init = False
 If Not SessionMe Then GoTo done
  
 describeBox.describeResultBox.Clear
 Dim tObject: For Each tObject In salesforce.EntityNames
    describeBox.describeResultBox.AddItem (tObject)
 Next tObject
 describeBox_init = True
 GoTo done

wild_error:
 MsgBox ("describeBox_init: wild_error")
 
done:
End Function
Function EntityNames()
If Not SessionMe Then GoTo done
  EntityNames = salesforce.EntityNames
done:
End Function
Public Function queryWizard_init()
  'On Error GoTo wild_error
  queryWizard_init = False
  If Not SessionMe Then GoTo done
  If Not set_Ranges Then GoTo done
  
  queryWiz.field.Clear
  queryWiz.field.AddItem "--Select One--"
  
  Dim labels(): ReDim labels(UBound(g_sfd.fields))
  Dim tObject, i As Integer
  For Each tObject In g_sfd.fields
    labels(i) = tObject.label:  i = i + 1:
  Next tObject
  
  Dim SwapValue, ix, jx   ' sort the list
  For ix = LBound(labels) To UBound(labels) - 1
    For jx = ix + 1 To UBound(labels)
      If labels(ix) > labels(jx) Then
        SwapValue = labels(ix): labels(ix) = labels(jx): labels(jx) = SwapValue
      End If
    Next
  Next

  '5.63 -
  ' going to add all fields, even if they only want to display some...
  Dim l: For Each l In labels
    queryWiz.field.AddItem l
  Next l

  
  queryWiz.field.Style = fmStyleDropDownList
  queryWiz.field.BoundColumn = 0
  queryWiz.field.ListIndex = 0

  queryWiz.operator.Clear
  queryWiz.operator.AddItem "equals"
  queryWiz.operator.AddItem "not equals"
  queryWiz.operator.AddItem "like"
  queryWiz.operator.AddItem "starts with"
  queryWiz.operator.AddItem "ends with"
  queryWiz.operator.AddItem "less than"
  queryWiz.operator.AddItem "less than or equal to"
  queryWiz.operator.AddItem "greater than"
  queryWiz.operator.AddItem "greater than or equal to"
  queryWiz.operator.AddItem "includes"
  queryWiz.operator.AddItem "excludes"
  queryWiz.operator.AddItem "regexp"
  queryWiz.operator.Style = fmStyleDropDownList
  queryWiz.operator.BoundColumn = 0
  queryWiz.operator.ListIndex = 0
  
  queryWizard_init = True
  GoTo done
wild_error:
  MsgBox ("queryWizard_init, wild error")
done:
  End Function
 
'
' find the column row intersection which has the ID we are looking for in this table
'
Function objectid(row, Optional quiet As Boolean) As String
  Dim t As Range: Set t = Application.Intersect(g_ids, ActiveSheet.Rows(row))
  On Error Resume Next
  objectid = t.value
  If Len(objectid) = 15 Then
    objectid = FixID(objectid)
  ElseIf LCase(objectid) = "new" Then
    objectid = objectid ' no change
  ElseIf Len(objectid) < 15 Then
    '  Debug.Print objectId
    If IsMissing(quiet) Then MsgBox "unrecognized object id >" & objectid & "<"
  End If
' Debug.Print "object id is " & objectId
End Function
'
' sanity check, is this a valid row to refresh
'
Function iSObject4Id(row) As Boolean
  iSObject4Id = False
  Dim t As Range: Set t = Application.Intersect(g_ids, ActiveSheet.Rows(row))
  If (t Is Nothing) Then GoTo done
  If IsError(t) Then GoTo done
  Select Case Len(t.value)
    Case 15, 18: iSObject4Id = True
  End Select
done:
End Function
Function default_query(g_table As Range)
  With g_table
    .Cells(1, 2).Interior.ColorIndex = 36
    .Cells(1, 3).Interior.ColorIndex = 36
    .Cells(1, 4).Interior.ColorIndex = 36
    .Cells(1, 2).value = recent_label(g_sfd) ' "System Modstamp" or ...
    .Cells(1, 3).value = ">"
    Dim MyDate: MyDate = Date
    .Cells(1, 4).value = MyDate - 7 ', "yyyy-mm-ddTHH:MM:SS.000Z")
  End With
End Function

'
' fetch assignment rule ids for options
'
Private Function getruleid(RuleType As String) As String
  getruleid = "no active rule found"
  Dim queryData As QueryResultSet4
  Application.StatusBar = "select data from AssignmentRule"
  Set queryData = salesforce.query("select id from AssignmentRule where " & _
    " active = TRUE and ruletype = '" & RuleType & "'")
  If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
  Dim so As SObject4
  For Each so In queryData
    getruleid = so.Item("Id").value
    Exit Function ' should be only one
  Next so
done:
End Function

'
' do the query and draw the rows we got back
' re-written at 5.26 to work-around a query bug in the COM dll. ( see below)
' first just get a list of ID's then pull the cols of each row using retireve,
' 5.27 add firstonly
' 5.92 optimize to stream large queries, don't store interim id list in a dict
'
Private Function query_draw(sels As String, where As String, outrow As Integer, _
                                                        firstonly As Boolean) As Integer
  Dim queryData As QueryResultSet4
  Application.StatusBar = Left("select data from " & g_objectType & " " & where, 128)
  '
  ' Debug.Print "select  " & sels
  ' Debug.Print "from " & g_objectType & " " & where
  
  'If (firstonly) Then ' this is not working to limit the batch size? 5.30
  '  Call salesforce.SetSoapHeader("QueryOptions", "batchSize", "1")
  '  End If
  On Error Resume Next
  Set queryData = salesforce.query("select id from " & g_objectType & " " & where)
  If (Not salesforce.Msg_and_Fail_if_APIError) Then
    If (firstonly) Then
        g_body.Cells(outrow, g_ids.Column - g_body.Column + 1).value = "#Err"
        outrow = outrow + 1
    End If
    GoTo done
    End If
  If queryData.Size = 0 Then
    If (firstonly) Then ' output something, like... "#N/F"
        g_body.Cells(outrow, g_ids.Column - g_body.Column + 1).value = "#N/F"
        outrow = outrow + 1
    End If
    GoTo done
  End If

  
  ' using a dict here fails, so we (fixed in 5.29)
  ' NEED TO write out as we read em...
  Application.ScreenUpdating = False
  Dim idlist() As String: ReDim idlist(maxBatchSize)
  Dim cnt%: cnt = 0: Dim remain%: remain = queryData.Size
  Dim ids As SObject4
  For Each ids In queryData
    idlist(cnt) = ids.Item("Id").value
    cnt = cnt + 1
    remain = IIf(firstonly, 0, remain - 1) ' 5.30
    '  Debug.Assert remain > 0
    If cnt > UBound(idlist) Or remain = 0 Then ' have a batch to do, or last one
      If (remain = 0) Then ReDim Preserve idlist(cnt - 1) ' now to pickup the tail end...
      Dim row As QueryResultSet4
      Set row = salesforce.Retrieve(sels, idlist, g_objectType)
      If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
    
      Dim so As SObject4: For Each so In row ' draw each of these rows,
        Call format_write_row(g_body, g_header, g_sfd, so, outrow)
        outrow = outrow + 1
        If outrow Mod maxBatchSize = 0 Then Call msg_user(outrow, queryData)
      Next so
      
      cnt = 0 ' reset the array counter after a batch completes
      DoEvents ' safe here ?
    End If
    If firstonly Then Exit For ' 5.30
  Next ids

  DoEvents
  Application.ScreenUpdating = True ' TODO manage this out side this function?
  
  Set ids = Nothing
done:
  query_draw = outrow
  
End Function

'
' the following function demonstrates a bug in the toolkit dll (which may be fixed by now...4/8/05)
' if many rows and cols are queried we sometimes get the
' description field back empty however, when pulling the same row using retireve it works
' that is to say,  the description field really has some data data ...
' switch to a method where all queries use retrieve to get cols (see above),
'
' function commented out on 12/7/04
'
'Private Function query_draw_old(sels As String, where As String, i As Integer) As Integer
'  Dim queryData As QueryResultSet4
'
'  Application.StatusBar = Left("select data from " & g_objectType & " " & where, 128)
'
' ' Debug.Print "select  " & sels
' ' Debug.Print "from " & g_objectType & " " & where
'   Set queryData = salesforce.query("select " & sels & " from " & g_objectType & " " & where)
'  ' Set queryData = salesforce.query("select * from " & g_objectType & " " & where)
'  If (Not salesforce.Msg_and_Fail_if_APIError) Then GoTo done
'  Application.ScreenUpdating = False
'
'  Dim so As SObject4: For Each so In queryData
'
'      Call format_write_row(g_body, g_header, g_sfd, so, i)
'
'      i = i + 1
'
'      If i Mod 50 = 0 Then ' provide progress updates for good user feel
'        Dim stmsg As String
'        stmsg = "Query : drawing query result " & i & " out of " & queryData.Size & " total returned"
'        If queryData.Size > 500 Then stmsg = stmsg & " -> press CTRL-BREAK to Cancel"
'        Application.StatusBar = stmsg
'
'        Application.ScreenUpdating = True
'        DoEvents ' this lets user break into long queries, like returning 54,000 contacts... oops
'        Application.ScreenUpdating = False
'      End If
'
'  Next so
'
'  query_draw = i
'done:
'End Function


Sub msg_user(i, queryData As QueryResultSet4)
  Dim stmsg As String
  stmsg = "Query : drawing query result " & i & " out of " & queryData.Size & " total returned"
  If queryData.Size > 500 Then stmsg = stmsg & " -> press CTRL-BREAK to Cancel"
  Application.StatusBar = stmsg
  Application.ScreenUpdating = True
  DoEvents ' this lets user break into long queries, like returning 54,000 contacts... oops
  Application.ScreenUpdating = False
End Sub
'
' returns a field from an table , given a name
'
Function findfield(sfd As SObject4, fld)
  Dim f%: For f = LBound(sfd.fields) To UBound(sfd.fields)
    If (sfd.fields(f).name = fld) Then
      Set findfield = sfd.fields(f)
      Exit Function
    End If
  Next f
End Function


'
' the ID has different names, like Account Id, so  use g_sfd
'
Public Function ensure_ID_select(theList As Object)
    Dim i%:  For i = 0 To theList.ListCount - 1
        ' given a label, return name, use slow version as labels are not cached yet
        If (sfFieldLabelToName_slow(g_sfd, theList.List(i)) = "Id") Then
            theList.Selected(i) = True
            Exit Function ' minor optimization...
        End If
    Next i
End Function

Function add_hyperlink(cel, soitem) ' trim down the value
    Dim link: link = Right(soitem.value, Len(soitem.value) - 5)
    link = Left(link, Len(link) - 5)
    Dim p: p = InStr(link, "_HL2_")
    If p < 1 Then Exit Function
    cel.value = Mid(link, p + 5)
    ActiveSheet.Hyperlinks.Add cel, Left(link, p - 1)
End Function
'
Function is_hyperlink(s)
    If (s.Type <> "string") Then Exit Function
    If (Left(s.value, 5) = "_HL1_") Then is_hyperlink = True
End Function


' 6.10
' code contributed by Erik Mittmeyer
' allows update only cells which are not hidden by filtering
'

Public Function AdjustFieldtype(vntValue As Variant, fld As SForceOfficeToolkitLib4.Field4)
'** this is toVBType slightly modified to work with Variants and also with early binding
'** parameter "fld". Also, I added a slightly more robust date handling and trimmed the strings
'** Actually, toVBType should work with variants too (I think)
    
    Select Case fld.Type
    Case "i4":
        AdjustFieldtype = Int(Val(vntValue))
    Case "double"
        AdjustFieldtype = Val(vntValue) ' normal case
    Case "percent"
        AdjustFieldtype = Val(vntValue)
        If Right(vntValue, 1) = "%" Then AdjustFieldtype = AdjustFieldtype / 100
    Case "datetime", "date"
        Dim vntUnInit As Variant
        AdjustFieldtype = IIf(IsDate(vntValue), CDate(vntValue), vntUnInit)
    Case "boolean":
        AdjustFieldtype = vntValue
    Case Else
      AdjustFieldtype = Trim$("" & vntValue)
    End Select

End Function

Sub sfUpdate_New()
    
    Const FunctionName = "sfUpdate_New" '** used in error handling at bottom of routine

    Dim xlSelection     As Excel.Range
    Dim xlRow           As Excel.Range
    Dim xlTempRow       As Excel.Range
    Dim xlColumn        As Excel.Range
    Dim xlCell          As Excel.Range
    
    Dim sd              As Scripting.Dictionary
    
    Dim cSobject        As SForceOfficeToolkitLib4.SObject4
    Dim rec()           As SForceOfficeToolkitLib4.SObject4
    Dim qr              As SForceOfficeToolkitLib4.QueryResultSet4
    
    Dim strArryCaps()   As String
    Dim strArryIDs()    As String
    Dim strArryCells()  As String
    Dim strFields       As String
    Dim strFieldName    As String
    Dim strAddrStart    As String
    Dim strAddrEnd      As String
    Dim strStatBarText  As String
    Dim strMsg          As String
    Dim strTitle        As String
    
    Dim vntArryVals()   As Variant
    Dim unInitVariant   As Variant
    
    Dim lngRows         As Long
    Dim lngCols         As Long
    Dim lngHiddenRows   As Long
    Dim lngHiddenCols   As Long
    
    Dim intFlags        As Integer
    Dim intBatchPointer As Integer
    Dim intJobPointer   As Integer
    Dim intFailedRows   As Integer
    Dim i As Integer, j As Integer
    
    Dim blnSkipHidden   As Boolean
    Dim blnGoAhead      As Boolean
    Dim blnNoLimits     As Boolean
    Dim blnGotCaptions  As Boolean

On Error GoTo Err_sfUpdate_New
    
    '// check for valid session or log in
    If Not SessionMe Then Err.Raise err_noSession   '// can be substituted with err_doNothing
    
    '// check global options
    blnGoAhead = QueryRegBool(GOAHEAD)
    blnNoLimits = QueryRegBool(NOLIMITS)
    'blnSkipHidden = QueryRegBool(SKIPHIDDEN)   '** new option setting?
    blnSkipHidden = True
        
    '// check the selected range
    Set xlSelection = Excel.Selection
    If xlSelection Is Nothing Then Err.Raise err_noSelection
    If xlSelection.Areas.Count > 1 Then Err.Raise err_tooManyAreas
    If Not set_Ranges Then Err.Raise err_doNothing  '// errors are already handled
    With xlSelection
        lngRows = .Rows.Count
        lngCols = .Columns.Count
    End With
    
    '// have we set the "Skip hidden fields" option?
    If blnSkipHidden = True Then
        '// count hidden rows
        For Each xlRow In xlSelection.Rows
            If xlRow.Hidden Then lngHiddenRows = lngHiddenRows + 1
        Next xlRow
        '// count hidden columns
        For Each xlColumn In xlSelection.Columns
            If xlColumn.Hidden Then lngHiddenCols = lngHiddenCols + 1
        Next xlColumn
    End If
    
    '// have we set the "Disregard reasonable limits" option?
    If Not blnNoLimits Then
        '// let's see if the selection is within the confines
        If lngRows - lngHiddenRows > maxRows Then
            '// maybe whole columns were selected; adjust selection to table body
            Intersect(xlSelection, g_body).Select
            Set xlSelection = Excel.Selection
            '// count rows again
            lngRows = xlSelection.Rows.Count
            lngHiddenRows = 0
            If blnSkipHidden = True Then
                For Each xlRow In xlSelection.Rows
                    If xlRow.Hidden Then lngHiddenRows = lngHiddenRows + 1
                Next xlRow
            End If
            If lngRows - lngHiddenRows > maxRows Then Err.Raise err_tooManyRows
        End If
        If lngCols - lngHiddenCols > maxCols Then
            '// maybe whole rows were selected; adjust selection to table body
            Intersect(xlSelection, g_body).Select
            Set xlSelection = Excel.Selection
            '// count columns again
            lngCols = xlSelection.Columns.Count
            lngHiddenCols = 0
            If blnSkipHidden = True Then
                For Each xlColumn In xlSelection.Columns
                    If xlColumn.Hidden Then lngHiddenCols = lngHiddenCols + 1
                Next xlColumn
            End If
            If lngCols - lngHiddenCols > maxCols Then Err.Raise err_tooManyCols
        End If
    End If
    
    '// have we set the "Skip warning dialogs" option?
    If Not blnGoAhead Then
        Dim s1 As String, s2 As String
        s1 = (lngRows - lngHiddenRows) & " row" & IIf(lngRows - lngHiddenRows = 1, "", "s")
        s2 = (lngCols - lngHiddenCols) & " column" & IIf(lngCols - lngHiddenCols = 1, "", "s")
        strMsg = "You are about to update " & s1 & " with " & s2 & "." & _
                vbCrLf & vbCrLf & "Do you want to proceed?"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then Err.Raise err_doNothing
    End If
    
    '// dimension the id and value arrays
    If lngRows - lngHiddenRows < maxBatchSize Then
        ReDim strArryIDs(lngRows - lngHiddenRows - 1)
        ReDim strArryCells(lngRows - lngHiddenRows - 1)
        ReDim vntArryVals(lngRows - lngHiddenRows - 1, lngCols - lngHiddenCols - 1)
    Else
        ReDim strArryIDs(maxBatchSize - 1)
        ReDim strArryCells(maxBatchSize - 1)
        ReDim vntArryVals(maxBatchSize - 1, lngCols - lngHiddenCols - 1)
    End If
    ReDim strArryCaps(lngCols - lngHiddenCols - 1)
    
    '// start processing
    strFields = "Id"
    For Each xlRow In xlSelection.Rows
        If intBatchPointer = 0 Then
            '// recreate for each new batch
            strStatBarText = intJobPointer + 1 & " -> " & _
                            IIf(lngRows - lngHiddenRows - intJobPointer > maxBatchSize, _
                            intJobPointer + maxBatchSize, lngRows - lngHiddenRows) & _
                            " of " & lngRows - lngHiddenRows
            Application.StatusBar = "Retrieving: " & strStatBarText
            DoEvents '// update status bar and allow interruption 6.13
        End If
        If Not (xlRow.Hidden And blnSkipHidden) Then
            '// fill elements of the batch; ignore hidden fields if required
            strArryIDs(intBatchPointer) = FixID(Intersect(g_ids, xlRow.EntireRow).value)
            i = 0
            For Each xlColumn In xlRow.Columns
                If Not (xlColumn.Hidden And blnSkipHidden) Then
                    If Len(strAddrStart) = 0 Then strAddrStart = xlColumn.Address
                    strAddrEnd = xlColumn.Address
                    xlColumn.Interior.ColorIndex = 36
                    If Not blnGotCaptions Then
                        '// fetch all captions, but only once in entire job
                        strArryCaps(i) = sfFieldLabelToName(g_sfd, _
                                        Intersect(g_header, xlColumn.EntireColumn).value)
                    End If
                    vntArryVals(intBatchPointer, i) = xlColumn.value 'it's actually only a cell
                    If Trim$(vntArryVals(intBatchPointer, i)) = "" Then
                        '// add fields with missing values only once per batch
                        If InStr(strFields, strArryCaps(i)) = 0 Then
                            strFields = strFields & ", " & strArryCaps(i)
                        End If
                    End If
                    i = i + 1
                End If
            Next xlColumn
            strArryCells(intBatchPointer) = strAddrStart & ":" & strAddrEnd
            strAddrStart = ""
            blnGotCaptions = True
            intBatchPointer = intBatchPointer + 1
            intJobPointer = intJobPointer + 1
        End If
        If intBatchPointer = maxBatchSize Or intJobPointer = lngRows - lngHiddenRows Then
            '// we've fetched a complete batch; let's retrieve the records
            Set qr = salesforce.Retrieve(strFields, strArryIDs, g_objectType)
            If (Not salesforce.Msg_and_Fail_if_APIError) Then Err.Raise err_doNothing
            Set sd = salesforce.ProcessQueryResults(qr)
            
            ReDim rec(UBound(strArryIDs))
            For i = LBound(strArryIDs) To UBound(strArryIDs)
                Set cSobject = sd.Item(strArryIDs(i))
                Debug.Assert IsObject(cSobject)
                
                For j = LBound(strArryCaps) To UBound(strArryCaps)
                    strFieldName = strArryCaps(j)
                    cSobject.Item(strFieldName).value = _
                        AdjustFieldtype(vntArryVals(i, j), cSobject.Item(strFieldName))
                    
                    If (cSobject.Item(strFieldName).value = "") Then
                      cSobject.Item(strFieldName).value = unInitVariant
                    End If
                    
                    If cSobject.Item(strFieldName).Type = "reference" Then
                      cSobject.Item(strFieldName).value = salesforce.ref_id( _
                          cSobject.Item(strFieldName).value, cSobject.Item(strFieldName).ReferenceTo)
                    End If
                Next j
                cSobject.Tag = i ' use the tag to store a row number on the worksheet
                Set rec(i) = cSobject
            Next i
            
            '** For i = LBound(rec) To UBound(rec)
            '**     For j = LBound(strArryCaps) To UBound(strArryCaps)
            '**         strMsg = strMsg & rec(i).Item(strArryCaps(j)).value & ","
            '**     Next j
            '**     strMsg = strMsg & vbCrLf
            '** Next i
            '** MsgBox strMsg
            
            
            '// now, let's do the update
            Application.StatusBar = "Updating: " & strStatBarText
            Call salesforce.DoUpdate(rec)

            For i = LBound(rec) To UBound(rec)
                '// check if the update was okay
                Set xlTempRow = Excel.ActiveSheet.Range(strArryCells(i))
                If rec(i).Error Then
                    xlTempRow.Interior.ColorIndex = 6
                    For Each xlCell In xlTempRow.Cells
                        If Not (xlCell.Comment Is Nothing) Then xlCell.Comment.Delete
                    Next xlCell
                    With xlTempRow.Cells(1, 1)
                        .AddComment
                        .Comment.Text "Row Update Failed" & vbLf & rec(i).ErrorMessage
                        .Comment.Shape.Height = 60
                    End With
                    intFailedRows = intFailedRows + 1
                Else
                    xlTempRow.Interior.ColorIndex = 0
                    For Each xlCell In xlTempRow.Cells
                        If Not (xlCell.Comment Is Nothing) Then xlCell.Comment.Delete
                    Next xlCell
                End If
            Next i
            
            '//reinitialize for next batch
            If lngRows - lngHiddenRows - intJobPointer < maxBatchSize _
              And lngRows - lngHiddenRows - intJobPointer <> 0 Then '6.13
                '// adjust array sizes, no need to preserve items
                ReDim strArryIDs(lngRows - lngHiddenRows - intJobPointer - 1)
                ReDim strArryCells(lngRows - lngHiddenRows - intJobPointer - 1)
                ReDim vntArryVals(lngRows - lngHiddenRows - intJobPointer - 1, lngCols - lngHiddenCols - 1)
            End If
            strFields = "Id"
            strAddrStart = ""
            strAddrEnd = ""
            intBatchPointer = 0
        End If
    Next xlRow
    
Exit_sfUpdate_New:
    '// finalizing code
    Application.StatusBar = Null
    If intFailedRows > 0 Then
        strMsg = intFailedRows & " of " & (lngRows - lngHiddenRows) & _
                " rows failed to update!" & vbCrLf & "Please check the comments" & _
                " in the first cell of each highlighted row for details"
        MsgBox strMsg, vbExclamation
    End If
    Exit Sub
    
Err_sfUpdate_New:
    strTitle = "Function aborted"
    intFlags = vbExclamation
    Select Case Err.Number
    Case err_doNothing
        strMsg = ""
    Case err_noSession
        strMsg = "There is no valid salesforce session."
    Case err_tooManyAreas
        strMsg = "You can't process multiple areas."
    Case err_tooManyRows
        strMsg = "You can't process more than " & maxRows & " rows."
    Case err_tooManyCols
        strMsg = "You can't process more than " & maxCols & " columns."
    Case err_noSelection 'will this ever happen? I wonder
        strMsg = "Nothing was selected"
'    Case err_noRows
'        strMsg = "No visible row in selection"
    Case err_noCols
        strMsg = "No visible column in selection"
    Case Else
        strMsg = Err.Description
        strTitle = "Error " & Err.Number & " in " & FunctionName & "()"
        intFlags = vbCritical
    End Select
    If Len(strMsg) > 0 Then
        MsgBox strMsg, intFlags, strTitle
    End If
    Resume Exit_sfUpdate_New '//good practice, just in case you want to clean up before leaving
  
End Sub
'*********************************************************************************************
'*********************************************************************************************


