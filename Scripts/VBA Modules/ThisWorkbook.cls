VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

With ShEmailRep
.ListObjects("From_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
.ListObjects("To_Date").DataBodyRange.Cells(1, 1).Value = "DD-MM-YYYY"
End With

Application.DisplayAlerts = False
Me.Connections("Query - Filtered Report").Refresh

Me.Save

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

'MsgBox "Manually Saving changes is disabled", vbExclamation, "Sorry"
'Cancel = True

End Sub

Private Sub Workbook_Open()

Dim shp As Shape

Application.DisplayAlerts = True

ShTracker.Protect UserInterfaceOnly:=True

With ShEmailRep

.Protect UserInterfaceOnly:=True, AllowUsingPivotTables:=True
    
    For Each shp In .Shapes
    
        If shp.Type = msoSlicer Then
        shp.Locked = False 'To enable the User to Toggle the slicers
        End If

    Next shp

End With
End Sub

