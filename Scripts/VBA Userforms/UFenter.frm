VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFenter 
   Caption         =   "Enter Details"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   OleObjectBlob   =   "UFenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtCancel_Click()

Unload Me

End Sub

Private Sub BtSave_Click()

Dim Cntrl As Control
Dim ReqInp As Boolean
Dim i As Integer

If Me.Caption = "Enter Details" Then 'Enter Details mode

    ReqInp = True 'by Default
    'Check for Required Textbox inputs
    For Each Cntrl In Me.Controls
     If TypeName(Cntrl) = "TextBox" Then
      If Cntrl.Tag = "Required" Then
       If Cntrl.Value = "" Then
        ReqInp = False
       End If 'Text
      End If 'Tag
     End If 'Typename
    Next Cntrl
    
    
    'Check for Required OptionButton inputs
    For i = 1 To 5
     If Me.Controls("ObY" & i).Value = False And Me.Controls("ObN" & i).Value = False Then
      ReqInp = False
     End If
    Next i
    
    
    'Check for Required Textbox for Opportunity/Angle field
    If CbOppty.Value = "Other" And TbOppty.Value = "" Then
    ReqInp = False
    LblWarn.Visible = True 'Pop up Other Opportunity Warning Label
    End If
    
    
    'Populating Warning on Required Inputs
    If ReqInp = False Then
    VBA.MsgBox "Please fill in all required fields"
    Exit Sub
    End If
    
    
    'Inputting all values in Tracker Table
    Dim TblTracker As ListObject
    Dim LastRow As Long
    Dim n As Long
    
    Set TblTracker = ShTracker.ListObjects("Tracker")
    LastRow = TblTracker.DataBodyRange.Rows.Count + 1 'Defining LastRow
    n = n + ShTracker.Range("XFD1").Value 'to make the Trxn no unique
    
    If TblTracker.DataBodyRange.Rows.Count < 2 And TblTracker.DataBodyRange(1, 1).Value = "" Then 'if there is only one last row left in Table
        TblTracker.DataBodyRange(1, 1).Value = 0.1
    End If
    
    With TblTracker
        'Column 1 to 6:
        .DataBodyRange(LastRow, 1) = "T" & n 'for unique Trxn no.
        .DataBodyRange(LastRow, 2) = VBA.Date
        .DataBodyRange(LastRow, 3) = Me.CbRepName.Value
        .DataBodyRange(LastRow, 4) = "=INDEX(Teams[Region],MATCH([@[Sales Agent]],Teams[Agent Name],0))"
        .DataBodyRange(LastRow, 5) = Me.TbCompet.Value
        .DataBodyRange(LastRow, 6) = Me.TbCust.Value
        .DataBodyRange(LastRow, 14) = "=IF(SUM(Tracker[@[Do we know the Executive Sponsor?]:[Unique Value Established?]])=5,1,1/SUM(Tracker[@[Do we know the Executive Sponsor?]:[Unique Value Established?]]))"
        .DataBodyRange(LastRow, 15) = "=IF(SUM(Tracker[@[Do we know the Executive Sponsor?]:[Unique Value Established?]])=5,""Predictable MRR"",""Deal behind Plan"")"
        
        'Input Opportunity Selected or Other:
         If Not Me.CbOppty.Value = "Other" Then
         .DataBodyRange(LastRow, 7) = Me.CbOppty.Value
         Else
         .DataBodyRange(LastRow, 7) = Me.TbOppty.Value
         End If
         
         'Current Date for Deal creation date:
        .DataBodyRange(LastRow, 8) = Me.TbDealAmt.Value
        
        'Checklist Yes/No:
         If Me.ObY1 = True Then
         .DataBodyRange(LastRow, 9) = 1
         Else
         .DataBodyRange(LastRow, 9) = 0
         End If
         
         If Me.ObY2 = True Then
         .DataBodyRange(LastRow, 10) = 1
         Else
         .DataBodyRange(LastRow, 10) = 0
         End If
         
         If Me.ObY3 = True Then
         .DataBodyRange(LastRow, 11) = 1
         Else
         .DataBodyRange(LastRow, 11) = 0
         End If
         
         If Me.ObY4 = True Then
         .DataBodyRange(LastRow, 12) = 1
         Else
         .DataBodyRange(LastRow, 12) = 0
         End If
         
         If Me.ObY5 = True Then
         .DataBodyRange(LastRow, 13) = 1
         Else
         .DataBodyRange(LastRow, 13) = 0
         End If
         
        'Column 8 (Projected Revenue):
        .DataBodyRange(LastRow, 16) = Me.TbProjAmt.Value
    End With
    
    If TblTracker.DataBodyRange(1, 1).Value = 0.1 Then
        Call Unprotect_Sheet
        TblTracker.DataBodyRange(1, 1).Delete
        Call Protect_Sheet
    End If
    
    Unload Me
    ThisWorkbook.Save
    

ElseIf Me.Caption = "Modify Details" Then 'Modify Details mode

    Call Modify_Details
    Unload Me
    ThisWorkbook.Save
    
End If


End Sub


Private Sub TbDealAmt_AfterUpdate()

If Not IsNumeric(Me.TbDealAmt.Value) Then
    MsgBox "Please enter a valid number", vbCritical
    Me.TbDealAmt.Value = ""
End If

End Sub

Private Sub TbOppty_AfterUpdate()

If Not Me.TbOppty.Value = "" Then
    Me.LblWarn.Visible = False
End If

End Sub

Private Sub TbProjAmt_AfterUpdate()

If Not IsNumeric(Me.TbProjAmt.Value) Then
    MsgBox "Please enter a valid number", vbCritical
    Me.TbProjAmt.Value = ""
End If

End Sub

Private Sub UserForm_Initialize()

Dim TblTeams As ListObject
Dim Tcell As Range

TbOppty.Enabled = False

'Populating Combobox for Sales Rep Name
Set TblTeams = ShTeams.ListObjects("Teams")

For Each Tcell In TblTeams.ListColumns(2).DataBodyRange
 If Not IsEmpty(Tcell.Value) Then
  CbRepName.AddItem Tcell.Value
 End If
Next Tcell

'Populating Opportunity/Angle dropdown
With CbOppty
.AddItem "Product Pricing"
.AddItem "Our USP"
.AddItem "Cust. Pain-Point"
.AddItem "Cust. Relationship"
.AddItem "Other"
End With

End Sub

Private Sub CbOppty_Change()

If Not CbOppty.Value = "Other" Then

 With TbOppty
  .Enabled = False
  .Value = ""
 End With
 
Else
TbOppty.Enabled = True
End If

End Sub
