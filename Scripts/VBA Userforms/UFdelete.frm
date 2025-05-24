VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFdelete 
   Caption         =   "Delete / Modify"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10665
   OleObjectBlob   =   "UFdelete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFdelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtDelCancel_Click()

Unload Me

End Sub

Private Sub BtGo_Click()

Dim i As Long

With Me

'If user Selects a Line and an Action (Modify/Delete):
    If .LbRecords.ListIndex <> -1 Then
        
        If Me.CbAction.Value = "Modify" Then
            Call Show_Modify_Form
        ElseIf Me.CbAction.Value = "Delete" Then
            Call Delete_Line_Item
            Exit Sub
        End If
    End If
    
'If user does not Select Anything:
    If .LbRecords.ListIndex = -1 Then
    
        MsgBox "Please Select a line to continue", vbCritical, "Attention" 'Asking user to select a line
    
    ElseIf .CbAction = "Select..." Then
    
        MsgBox "Please Select an Action to continue", vbCritical, "Attention" 'Asking user to select an Action
    
    End If
    
End With
    
End Sub


Private Sub UserForm_Initialize()

Dim TblTracker As ListObject

Set TblTracker = ShTracker.ListObjects("Tracker")

With UFdelete

.LbRecords.ColumnCount = TblTracker.ListColumns.Count - 1
.LbRecords.RowSource = "Tracker[#Data]"
.LbRecords.ColumnHeads = True
.CbAction.AddItem "Delete"
.CbAction.AddItem "Modify"
.CbAction.Value = "Select..."

End With


End Sub

