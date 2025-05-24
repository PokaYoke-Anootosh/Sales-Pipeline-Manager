Attribute VB_Name = "Backend"
Option Explicit

Sub UFenter_Form()

With UFenter
.Caption = "Enter Details"
.Show
End With

End Sub

Sub UFdelete_Form()

UFdelete.Show

End Sub


Sub Show_Modify_Form()

Dim TblTracker As ListObject
Dim TrxnNo As String
Dim i As Long
Dim m As Integer

Set TblTracker = ShTracker.ListObjects("Tracker")
TrxnNo = UFdelete.LbRecords.List(UFdelete.LbRecords.ListIndex, 0) 'Column 0 = Trxn no.


For i = 1 To TblTracker.ListRows.Count
    
    If TblTracker.DataBodyRange.Cells(i, 1).Value = TrxnNo Then
        Exit For 'To keep the value of i preserved
    End If

Next i


With UFenter

.Caption = "Modify Details"

'Populate previously selected Textbox field values:
.CbRepName.Value = TblTracker.DataBodyRange.Cells(i, 3).Value
.TbCompet.Value = TblTracker.DataBodyRange.Cells(i, 5).Value
.TbCust.Value = TblTracker.DataBodyRange.Cells(i, 6).Value
.TbDealAmt.Value = TblTracker.DataBodyRange.Cells(i, 8).Value
.TbProjAmt.Value = TblTracker.DataBodyRange.Cells(i, 16).Value


'Populating Opportunity field options previously selected:
For m = 0 To UFenter.CbOppty.ListCount - 1
    If UFenter.CbOppty.List(m) = TblTracker.DataBodyRange.Cells(i, 7).Value Then
    .CbOppty.Value = TblTracker.DataBodyRange.Cells(i, 7).Value
    Exit For
    End If
Next m
    
If .CbOppty.Value = "" Then 'if cboopty options have not matched to provious values
    .CbOppty.Value = "Other"
    .TbOppty.Value = TblTracker.DataBodyRange.Cells(i, 7).Value
End If
    
       
'Populating Checkbox options previously selected:
    If TblTracker.DataBodyRange.Cells(i, 9).Value = 1 Then
    .ObY1 = True
    Else
    .ObN1 = True
    End If
    
    If TblTracker.DataBodyRange.Cells(i, 10).Value = 1 Then
    .ObY2 = True
    Else
    .ObN2 = True
    End If
    
    If TblTracker.DataBodyRange.Cells(i, 11).Value = 1 Then
    .ObY3 = True
    Else
    .ObN3 = True
    End If
    
    If TblTracker.DataBodyRange.Cells(i, 12).Value = 1 Then
    .ObY4 = True
    Else
    .ObN4 = True
    End If
    
    If TblTracker.DataBodyRange.Cells(i, 13).Value = 1 Then
    .ObY5 = True
    Else
    .ObN5 = True
    End If

.Show 'Show the Entry Userform

End With

End Sub

Sub Modify_Details()

Dim TrxNo As String
Dim TblTracker As ListObject
Dim i As Long

TrxNo = UFdelete.LbRecords.List(UFdelete.LbRecords.ListIndex, 0)
Set TblTracker = ShTracker.ListObjects("Tracker")

Call Unprotect_Sheet

With TblTracker

For i = 1 To .DataBodyRange.Rows.Count

    If .DataBodyRange.Cells(i, 1).Value = TrxNo Then
        Exit For
    End If

Next i

'Filling details from Textboxes & Comboboxes:
.DataBodyRange.Cells(i, 3).Value = UFenter.CbRepName.Value
.DataBodyRange.Cells(i, 5).Value = UFenter.TbCompet.Value
.DataBodyRange.Cells(i, 6).Value = UFenter.TbCust.Value

If .DataBodyRange.Cells(i, 7).Value = "Other" Then
.DataBodyRange.Cells(i, 7).Value = UFenter.TbOppty.Value
Else
.DataBodyRange.Cells(i, 7).Value = UFenter.CbOppty.Value
End If

.DataBodyRange.Cells(i, 8).Value = UFenter.TbDealAmt.Value
.DataBodyRange.Cells(i, 16).Value = UFenter.TbProjAmt.Value
.DataBodyRange.Cells(i, 17).Value = VBA.Date


'Checking Options from Checkbox inputs:
    If UFenter.ObY1 Then
        .DataBodyRange.Cells(i, 9).Value = 1
    Else
        .DataBodyRange.Cells(i, 9).Value = 0
    End If
    
    If UFenter.ObY2 Then
        .DataBodyRange.Cells(i, 10).Value = 1
    Else
        .DataBodyRange.Cells(i, 10).Value = 0
    End If
    
    If UFenter.ObY3 Then
        .DataBodyRange.Cells(i, 11).Value = 1
    Else
        .DataBodyRange.Cells(i, 11).Value = 0
    End If
    
    If UFenter.ObY4 Then
        .DataBodyRange.Cells(i, 12).Value = 1
    Else
        .DataBodyRange.Cells(i, 12).Value = 0
    End If
    
    If UFenter.ObY5 Then
        .DataBodyRange.Cells(i, 13).Value = 1
    Else
        .DataBodyRange.Cells(i, 13).Value = 0
    End If

End With

MsgBox "Your changes have been saved", vbInformation, "Success"

Call Protect_Sheet

End Sub

Sub Delete_Line_Item()

Dim TblTracker As ListObject
Dim TrxNo As String
Dim YesNo As VbMsgBoxResult
Dim i As Long

Set TblTracker = ShTracker.ListObjects("Tracker")
TrxNo = UFdelete.LbRecords.List(UFdelete.LbRecords.ListIndex, 0)
YesNo = MsgBox("Are you sure you want to delete this item?", vbYesNo, "Confirmation")


If YesNo = vbNo Then 'To take confirmation before deleting

    Exit Sub

Else

    Call Unprotect_Sheet
    
    For i = 1 To TblTracker.DataBodyRange.Rows.Count
        If TblTracker.DataBodyRange.Cells(i, 1).Value = TrxNo Then
            Exit For
        End If
    Next i
    
    With TblTracker
    
        If .DataBodyRange.Rows.Count < 2 And .DataBodyRange.Cells(i, 1).Value = "" Then
            
            MsgBox "No Records found to delete", vbExclamation, "Oops"
            
            Exit Sub
        
        ElseIf .DataBodyRange.Rows.Count < 2 Then
        
            .ListRows(i).Range.ClearContents
            
            MsgBox "The line item you selected has been deleted off the records", vbInformation, "Success"
            
        Else
                .ListRows(i).Delete
            
            MsgBox "The line item you selected has been deleted off the records", vbInformation, "Success"
        
        End If
    
    End With
    
    Call Protect_Sheet

End If

End Sub

Sub Protect_Sheet()

ShTracker.Protect UserInterfaceOnly:=True

End Sub

Sub Unprotect_Sheet()

ShTracker.Unprotect

End Sub

Sub Test_bed()

Debug.Print

End Sub

Sub Downld_Report()

Dim MyPvt As PivotTable
Dim FilePickr As FileDialog
Dim filepath As String
Dim TempWb As Workbook
Dim ShTemp As Worksheet
Dim shTemp2 As Worksheet

Set MyPvt = ShEmailRep.PivotTables("Email_Report")

If ShEmailRep.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY" Or _
ShEmailRep.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY" Then 'if no date range is selected

    MsgBox "Please enter a valid date before proceeding with donwload/export", vbExclamation, "Sorry"
    
Else
    
    Application.ScreenUpdating = False 'Turn off ScreenUpdating to optimize

    'Copying visible cells from Pivot Table Email_Report
    MyPvt.TableRange1.SpecialCells(xlCellTypeVisible).Copy
    
    
    'Setting up my attachment Workbook and Worksheet
    Set TempWb = Workbooks.Add
    Set ShTemp = TempWb.Sheets(1)
    ShTemp.Name = "Sales Pipeline Report"
    
    
    'Copying Pivot data in New Workbook ShTemp
    MyPvt.TableRange1.Copy
    
    With ShTemp.Range("A1")
    .PasteSpecial Paste:=xlPasteValues
    .PasteSpecial Paste:=xlPasteFormats
    End With
    
    Application.CutCopyMode = False
    
    
    'Saving the attachment sheet
    Set FilePickr = Application.FileDialog(msoFileDialogSaveAs)
    
    With FilePickr
    .Title = "Save Report"
    
    .InitialFileName = "Sales_Pipeline_Report_" & VBA.Format(ShEmailRep.Range("A3"), "yyyy-mm-dd") _
    & "_to_" & VBA.Format(ShEmailRep.Range("C3"), "yyyy-mm-dd") & ".xlsx"
    
    .FilterIndex = 1 'To make Excel Workbook/.xlsx format the first category in filedialog file-filter
    
        If .Show <> -1 Then 'i.e if user clicks the Cancel button (value of .Show will be -1)
        MsgBox "Export Cancelled", vbInformation
        TempWb.Close SaveChanges:=False
        Exit Sub
        End If
    
    filepath = .SelectedItems(1)
    End With
    
    With TempWb
    .SaveAs filepath
    .Close
    End With

    Application.ScreenUpdating = True 'Turn ScreenUpdating back on

End If

End Sub

