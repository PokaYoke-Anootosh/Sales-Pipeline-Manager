VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShEmailRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Pop_Date_InputBox()

Dim selectedDate As Variant

If Me.Range("XFD1").Value = "From" Then

selectedDate = Application.InputBox("Date From:", Title:="Pick a date (DD-MM-YYYY)", Type:=2)
    If IsDate(selectedDate) Then
        Me.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = CDate(selectedDate)
    ElseIf selectedDate = "" Then
        Me.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
    Else
        MsgBox "Please enter a Date value", vbExclamation, "Value Error"
        Me.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
    End If
    
ElseIf Me.Range("XFD1").Value = "To" Then

selectedDate = Application.InputBox("Date Till:", Title:="Pick a date (DD-MM-YYYY)", Type:=2)
    If IsDate(selectedDate) Then
        Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = CDate(selectedDate)
    ElseIf selectedDate = "" Then
        Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
    Else
        MsgBox "Please enter a Date value", vbExclamation, "Value Error"
        Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
    End If
End If

End Sub

Private Sub CbCalendar_FROM_Click()

Me.Range("XFD1").Value = "From"

Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"

Call Pop_Date_InputBox

End Sub

Private Sub CbCalendar_TO_Click()

Me.Range("XFD1").Value = "To"

Call Pop_Date_InputBox

If Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY" Then 'dont run if To_Date table is not showing a date value
    Exit Sub
Else
    Call Refresh_Email_Report
End If

End Sub

Private Sub Worksheet_Activate()

Me.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
Me.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"

End Sub

Private Sub Refresh_Email_Report()

ThisWorkbook.Connections("Query - Filtered Report").Refresh
ShEmailRep.Range("A12").Select

End Sub
