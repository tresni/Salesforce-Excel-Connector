Attribute VB_Name = "AutoExec"
Option Explicit
'
' Menus and dialog init
' comments and change history by Ron Hess
'
'
Private Const IDS_TOOLBAR_NAME As String = "Force.com Connector"
Private Const Ver As String = "16.03 - 7/14/2011"
' 16.03 - custom fields check should exclude required, those are already shown , avoids dups
' 16.02 - put the menu in the menu bar, remove ref to refedit.DLL (missing in win 7?)
' 9.01 - link with the office toolkit found in latest spring 09 outlook edition (13.0 api)
' 6.16 - insure that incorrect password msg is posted when bad passwd is sent in
' 6.15 - back out 6.14, won't clear err status until logout
' 6.14 - better msg when api server is down, well not much better...
' 6.13 - minor fix from erik for 6.11 changed routine
' 6.12 - return a number found from sub sfQuery(), thx to David A. Korba
' 6.11 - fix for int into sfQueryValueFormat(), screening on an integer value. attachment // Body Length // > // 300000
' 6.10 - large routine from Erik, allows updates to skip hidden cells , set by option
' 6.09 - show hyperlink formulas as links in excel, spiffy
' 6.08 - fix lost decimals in europe
' 6.07 - back out 6.04, was removing wordwrap formating on textareas, use ClearContents
' 6.06 - add fixidsel() to fixup only the cells in the selection (15char ->18)
' 6.05 - prompt to download the v3.0 toolkit if it's not found
' 6.04 - when clearing the area below a query, really clear it, comments also
' 6.03 - backout 5.61, the v3.0 dll fixes the search() bug
' 6.02 - finaly a fix for dates getting americanised, thanks to mikeol
' 6.01 - port to V3.0 toolkit uses SObject4, ...
' 5.67 - scot s points out another issue with percent,also dont add ".0" if double has a decimal already
' 5.66 - use Application.International(xlDateOrder) to format dates
' 5.65 - trim strings used in the filter cell before placing into a query
' 5.64 - replace + with & in string concat
' 5.63 - more wizard steps to select fields, from a valued contributor at odyssey-group
' 5.62 - if we hit an invalid sessionid from timeout, hit the logout function
' 5.61 - rush out a workaround for toolkit.Search() memfault, reported to sf.com
' 5.60 - add exception code to error message dialog
' 5.59 - work on session timeout retry, not done yet
' 5.58 - add optional parm to sfsearch()
' 5.57 - add api name to comments returned by describe
' 5.56 - fix a ON join problem with "and and" showing up
' 5.55 - code to support Contact and Account dicts,
'   this controled optionaly by a flag for performance concerns, default off
' 5.51 - ref_id will now look in both group and user dicts if ref_to is Group
' 5.49 - allow query to work with "empty" dates, generate soql -> "datefield = null"
' 5.48 - bug in queryAll when slient is turned on, thanks Barry K. clean up some comments
' 5.47 - fix percentages, in excel are of the form "0.55" == 55% in sf "0.55" is < 1%
' 5.46 - fixes role names in query, query now uses ref_id, comment out aliasfield
' 5.45 - take out the query row check for valid id in the first row, not helping
'        working on allowing #N/F to appear in the ID col, and we skip these
' 5.44 - bug fix with record id as a name on insert
' 5.43 - allow update to run past failed rows, mark failures with cel comments
' 5.42 - Ezra provides upgrades query wiz, tweak to not delete cells
' 5.41 - allow user default browser for help popup, thanks to a sharp community member
' 5.40 - correct ids returned when a group name is empty, broken in 5.35
' 5.37 - fix 5.29 when ref is for a custom object ref_id will restore incoming value
' 5.36 - batch the create call, i learned how to do this... yeah!
' 5.35 - add dict for groups,roles,profiles
' 5.34 - fix starts with, was not working for a list of starts with, thx to Timothy Bouscal
'        pop a good message on failed passwd, was just showing an empty box
' 5.33 - add starts with, ends with and regexp as operators
' 5.32 - on insert allow strings for recordtypes,
' 5.31 - on insert allow strings for usernames, like done for query in 5.29
' 5.30 - setsoap batchsize does not appear to limit the query, adjust for ON join
' 5.29 - allow usernames to be returned inplace of ID's (by default now),
'   discovered and addressed large query issue introduced in 5.26
' 5.28 - fix sessionme to avoid re-logins, set batchSize 1 when doing ON join
' 5.27 - add ON join, which returns the first record matching given a range
'   this tries to get the returned data lined up with the same row in an
'   nearby table
' 5.26 - re-write query_draw to pull data in chunks with multiple calls to
'   retreve, solves some missing data in description fields with larger queries
' 5.25 - point to a help file on sourceforge
' 5.24 - add option to allow raw api names in cols, fix queryall sort order
' 5.23 - check for picklist and like operator, warn user
' 5.22 - add link to online help
' 5.21 - sort the areas when running all on a sheet,
'   escape tick ->'< in query values
' 5.20 - add IN operator for more join support
' 5.19 - found a bug with very large messages sent to the status bar, throws an error
'   will truncate this case
' 5.18 - queryreg needs to look at the error code
' 5.17 - lead and case assignments in options dialog setsoapheader to trigger rules
' 5.16 - add option to skip limits on update & insert
' 5.15 - adding options dialog for 5.0 server URL mostly,
'  new code to perform a "refresh" , which is the same as -> Query All Rows for current table
'   useful in creating "left join" using two regions of cells
'   test object ID to ensure it is of the type specified by the table name
'   try to format currency and text areas properly in output of sfQuery
' 5.14 - fix setting dates to empty, need to recognize this special case
' 5.13 - add multiple query pulldown menu, used for running testplan
' 5.12 - support TODAY , TODAY - n, TODAY + n in date fields for sfQuery()
' 5.11 - sfQuery() fix case of value passed in is an empty cel, was skiping this case
' 5.10 - add queryeachworksheet functionality, for test plan
'   fix insert to look at Createable not Updateable when adding fields
'   describe will now decorate all fields with their type in a comment
' 5.09 - use getservertime to check for a valid session, not getuserid
'   dont cache user name info, just fetch it each time in sfuserid() and sfusername()
' 5.08 - verify the session to avoid using one with and expired or invalid sid
'   sort field list in wizard, more wizard testing, more helpful messages
' 5.07 - check for active sheet before catching wild error
' 5.06 - query wizard appears in menu as the new users first choice
' 5.05 - describe now orders std fields first then custom, then read-only
'   fix insertRow to not insert into read-only fields
'   decorate picklist columns with comments showing the pick values
'   crude query wizard added
' 5.04 - add support for comma sep lists in value field of query ~ sfQuery()
'   turn off screen updating while selecting all areas in sfQueryAll(), add a message
' 5.03 - more testing on SOAP 5.0,
'   check for com object and provide a message if missing
' 4.62 - multipick type needs "(...)" construction in sfQuery()
' 4.61 - fix advanced option to allow prerelease endpoints to work correctly
'   much better caching of sessions, avoids extra login calls
'   set VBA References to SForceOfficeToolkit v2.0
' 4.60 - fixed the required fields when building Contracts (allow autonumber)
' 4.59 - allow combo box field type in sfQuery,
'   should probably just default unknown types to string...
'

'
' each time we launch the add-in the is called
'
Sub Auto_Open()
    On Error Resume Next
  '  CommandBars(IDS_TOOLBAR_NAME).Delete  ' not neeeded each open
   ' CommandBars("Salesforce").Delete ' remove old toolbar
       
    Dim ourToolbar As CommandBar
    Set ourToolbar = CommandBars(IDS_TOOLBAR_NAME)
    If Application.CommandBars(IDS_TOOLBAR_NAME).Controls(1).Tag <> Ver Then
        CommandBars(IDS_TOOLBAR_NAME).Delete ' saw new version
        ' so we rebuild the toolbar incase new buttons have been added...
        Set ourToolbar = Nothing
    End If
    
    If ourToolbar Is Nothing Then
        ' 5.03 - check it here, before we get a copy of the src code splashed on screen
        ' before we make a toolbar, check that the SForceOffice Toolkit is available
        Dim sfot As Object
        Set sfot = New SForceSession4 ' if this fails we dont have the new toolkit
        If sfot Is Nothing Then
            
            GoTo done
        End If
        sfot = Nothing ' done with our test, dont need this here
        Err.Clear
        Set ourToolbar = CommandBars.Add(name:=IDS_TOOLBAR_NAME, _
            Position:=msoBarTop _
            )
        ourToolbar.Visible = True
        
        CreateCustomMenuWithSubMenus (IDS_TOOLBAR_NAME)
    End If
done:
End Sub

Sub CreateCustomMenuWithSubMenus(tbname As String)
    Dim mMain As CommandBarPopup, mSub As CommandBarPopup
    Set mMain = Application.CommandBars(tbname).Controls.Add( _
        msoControlPopup, 1, , , False)
    With mMain
        .Caption = "Force.com &Connector"
        .Tag = Ver  ' put the version somewhere
        .TooltipText = "Connect to and exchange data with Salesforce.com"
        .HelpContextID = 2
        .HelpFile = "sforce_connect.chm"
    End With

    With mMain.Controls.Add(Type:=msoControlButton)
        .Caption = "Table Query &Wizard"
        .OnAction = "sfDescribeAndQuery"
        .TooltipText = "Select sforce object, describe and query its contents"
        .Style = msoButtonIconAndCaption
        .FaceId = 581
    End With
    
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 2892
        .OnAction = "sfUpdate"
        .Caption = "&Update Selected Cells"
        .TooltipText = "send an Update call to salesforce.com passing the values in the selected cells"
    End With
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 539
         .OnAction = "sfInsertRow"
        .Caption = "&Insert Selected Rows"
        .TooltipText = "Insert (new) one row of data from Salesforce.com"
    End With
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 136
         .OnAction = "sfQueryRow"
        .Caption = "Query Selected &Rows"
        .TooltipText = "Query one or more rows (selected) of data from Salesforce.com"
    End With
 
   With mMain.Controls.Add(Type:=msoControlButton)
   .BeginGroup = True
        .Caption = "&Describe Sforce Object"
        .OnAction = "sfDescribe"
        .TooltipText = "Describe valid columns for the sepecified Salesforce object"
        .Style = msoButtonIconAndCaption
        .FaceId = 133
    End With
    
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 459
        .OnAction = "sfQuery"
        .TooltipText = "Run the Query in the first row of the current region, return table data from Salesforce"
        .Caption = "&Query Table Data"
    End With
    
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 348
        .OnAction = "sfDelete"
        .Caption = "Delete Objects"
        ' .enabled = false ' this flag would be nice to maintain
    End With
        
    ' add quick start submenus...
    Set mSub = mMain
    CreateSubMenu1 mSub

'    With mMain.Controls.Add(Type:=msoControlButton)
'    End With
    
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 348
        .OnAction = "sfLogout"
        .Caption = "&Logout Session"
        ' .enabled = false ' this flag would be nice to maintain
    End With
    
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 607
        .Enabled = False
        .Caption = "no user name"
        .Tag = "username"
    End With
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .BeginGroup = True
        .FaceId = 3116
        .OnAction = "sfOptions"
        .Caption = "Options"
      .TooltipText = "Set Default Server URL \& other options"
    End With
    With mMain.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .BeginGroup = False
        .FaceId = 345
        .OnAction = "sfAbout"
        .Caption = "Help on sforce-Excel Add-In"
    End With
    
    Set mSub = Nothing
    Set mMain = Nothing
End Sub
Public Function displayUserName(nam As String)
On Error Resume Next
    Dim userbutton As CommandBarButton
    Set userbutton = CommandBars(IDS_TOOLBAR_NAME).FindControl( _
        msoControlButton, , "username", 1, 1)
    userbutton.Caption = nam ' return the button
    displayUserName = True
End Function

'
'
'
Sub CreateSubMenu1(InputCtrl As CommandBarPopup)
Dim SubMenu As CommandBarPopup
    ' create the new submenu
    Set SubMenu = InputCtrl.Controls.Add(Type:=msoControlPopup)
    With SubMenu ' add the menu caption
        .BeginGroup = True
        .Caption = "&Multiple Queries"
        .Tag = "MySubMenuTag"
    End With

    With SubMenu.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 459
        .OnAction = "sfQueryAll"
        .Caption = "Run &Each Query on &Current Sheet"
        .TooltipText = "visit all tables on this worksheet, run the Query contained in each one"
    End With
    With SubMenu.Controls.Add(Type:=msoControlButton)
        .Style = msoButtonIconAndCaption
        .FaceId = 133
        .OnAction = "sfQueryAllSheets"
        .Caption = "Run Each Query on &All Sheets"
        .TooltipText = "visit all tables on Every worksheet, run the Queries contained in each one"
    End With
done:
    Set InputCtrl = SubMenu ' return the new submenu control to the calling procedure
    Set SubMenu = Nothing
End Sub

Sub DescribeAccount()
    Selection.value = "account"
    Call sfDescribe
End Sub
Sub DescribeOpp()
    Selection.value = "opportunity"
    Call sfDescribe
End Sub
Sub DescribeContact()
    Selection.value = "contact"
    Call sfDescribe
End Sub
Sub DescribeCase()
    Selection.value = "case"
    Call sfDescribe
End Sub
Sub DescribeAll()
    Selection.value = "All Sforce Entities"
    Call sfDescribe
End Sub

' walk thru all active regions on the worksheet
' call sfQuery on each region
' select the first non blank cell
' could be very handy to build a join or pivot
Sub sfQueryAll(Optional silent)
  Dim old_pos As Range: Set old_pos = ActiveCell
  Dim used As Range
  Set used = Range("A1", ActiveCell.SpecialCells(xlLastCell).Address)
  If used.Count = 1 Then GoTo done ' nothing on this sheet
  used.Find("*", LookIn:=xlValues).CurrentRegion.Select
  
  Dim r1, c As Range: Set r1 = Selection
  Application.ScreenUpdating = False ' (5.04) quite a show if we dont do this
  For Each c In used.Cells ' include non blank in this range
  If c.Text <> "" Then ' this is not exactly fast, but it works
    Union(r1, c.CurrentRegion).Select
    Set r1 = Selection
  End If
  Next c
  Application.ScreenUpdating = True
  
  If Not IsMissing(silent) Then GoTo ready
  
  Dim msg ' warn the user that they are writing to their worksheet 5.04
  msg = "You are about to QUERY: " & CStr(Selection.Areas.Count) & _
    " tables in the current worksheet, this will overwrite the data in each table"
  If (MsgBox(msg, vbApplicationModal + vbOKCancel + vbExclamation + vbDefaultButton1, _
       "-- Ready to Query Salesforce.com --") = vbCancel) Then GoTo done
 
ready:
 ' sometimes the areas are not sorted, to be consistent, sort them
 ' start by saving the address of each area 5.21
 Dim myareas() As String: ReDim myareas(Selection.Areas.Count - 1)
 Dim idx%: idx = 0: For Each r1 In Selection.Areas
    myareas(idx) = r1.AddressLocal: idx = idx + 1
 Next r1
 
 ' sort my areas, by their address, puts the left most range first 5.21
 Dim SwapValue, ix, jx   ' sort the list
 For ix = LBound(myareas) To UBound(myareas) - 1
    For jx = ix + 1 To UBound(myareas)
     Dim ixr As Range: Set ixr = Range(myareas(ix))
     Dim jxr As Range: Set jxr = Range(myareas(jx))
     If ixr.Column > jxr.Column Then ' compare addresses note, $B$1 < $AB$1
      SwapValue = myareas(ix): myareas(ix) = myareas(jx): myareas(jx) = SwapValue
      End If
    Next
 Next
  
 'For ix = LBound(myareas) To UBound(myareas): Debug.Print myareas(ix): Next ix

  Dim entnames: entnames = s_force.EntityNames
  For ix = LBound(myareas) To UBound(myareas) ' 5.21
    'Debug.Print myareas(ix)
    Dim ar As Range: Set ar = Range(myareas(ix)) ' make the string back into a range
    ' queryall needs to be smart enough to run on workbooks with other
    ' information (not queries) present,
    ' so we check cell(1,1) for a valid entity
    Dim e: For Each e In entnames
      If (LCase(e) = LCase(ar.Cells(1, 1).value)) Then
        ar.Select: Call sfQuery ' what could be easier...
        End If
      Next e
  Next ix
  
done:
  old_pos.Select
End Sub
'
' do the above for each worksheet, slick..
' used to run test plan 5.10
'
Public Function sfQueryAllSheets()
  Dim sel As Worksheet:   Set sel = Application.ActiveSheet
  Dim rng As Range:       Set rng = Selection
  
  Dim msg ' warn the user that they are writing to their worksheet 5.04
  msg = "You are about to re-run --- EACH QUERY on ALL Worksheets --- in this workbook: " & _
    vbCrLf & "this will overwrite the existing data in all table"
  If (MsgBox(msg, vbApplicationModal + vbOKCancel + vbExclamation + vbDefaultButton1, _
       "-- Ready to Query Salesforce.com --") = vbCancel) Then GoTo done
       
  Dim s As Worksheet: For Each s In Application.Worksheets
    s.Activate: Call sfQueryAll(True)
  Next s
  
done:
  sel.Activate: rng.Select ' jump back to where we started
End Function
'
' easier to support if this is a clean slate 9.1
'
Sub Auto_Close()
    On Error Resume Next
    CommandBars(IDS_TOOLBAR_NAME).Delete
End Sub
Sub sfOptions()
     options.Show
End Sub

Sub sfAbout()
     aboutBox.Show
End Sub

Sub sfDescribeAndQuery()
On Error GoTo nosheet
  Dim used As Range: Set used = ActiveCell
  If s_force.describeBox_init = True Then step1.Show
  GoTo done
nosheet:
  MsgBox "Oops, Could not find an active Worksheet"
done:
End Sub

Public Function ver_str() As String
 ver_str = Ver
End Function

