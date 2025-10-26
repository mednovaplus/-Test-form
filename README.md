# -Test-form
My first inventory
Option Explicit
Dim PatRow As Long, PatCol As Long, LastRow As Long, LastResultRow As Long

Sub Patient_RefreshList()
    MedHist.Range("C4:D99").ClearContents
    MedHist.Shapes("EditPatientBtn").Visible = msoFalse
    With Patients
        .Range("U3:Z9999").ClearContents
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        .Range("A3:K" & LastRow).AdvancedFilter xlFilterCopy, , CopyToRange:=.Range("U2:X2"), Unique:=True
        LastResultRow = .Range("X99999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        .Range("Y3:Z" & LastResultRow).Formula = .Range("Y1:Z1").Formula
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=Patients.Range("Y3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
            .SetRange Patients.Range("U3:Z" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
        End With
        
NoSort:
        MedHist.Range("C4:D" & LastResultRow + 1).Value = .Range("X3:Y" & LastResultRow).Value
    End With
End Sub

Sub Patient_Edit()
    If MedHist.Range("B3").Value = "" Then Exit Sub
    PatRow = MedHist.Range("B3").Value
    With PatFrm
        For PatCol = 2 To 10
            .Controls("Field" & PatCol - 1).Value = Patients.Cells(PatRow, PatCol).Value
        Next PatCol
        .PatRow.Value = PatRow
        .Show
    End With
End Sub

Sub Patient_AddNew()
    PatFrm.Show
End Sub

Sub Patient_SaveUpdate()
    With PatFrm
        If .Field1.Value = Empty Then
            MsgBox "Please make sure to add in an Patient Name before saving"
            Exit Sub
        End If
        If .PatRow.Value = Empty Then
            PatRow = Patients.Range("A99999").End(xlUp).Row + 1
            Patients.Range("A" & PatRow).Value = MedHist.Range("B10").Value 'Next Patient ID
            Patients.Range("O" & PatRow).Value = "=Row()"
        Else
            PatRow = PatRow
        End If
        For PatCol = 2 To 10
            Patients.Cells(PatRow, PatCol).Value = .Controls("Field" & PatCol - 1).Value
        Next PatCol
        Unload PatFrm
        Patient_RefreshList
    End With
End Sub

Sub Patient_Delete()
    If MsgBox("Are you sure you want to delete this Patient?", vbYesNo, "Delete Patient") = vbNo Then Exit Sub
    If PatFrm.PatRow.Value = Empty Then GoTo NotSaved
    PatRow = PatFrm.PatRow.Value
    Patients.Range(PatRow & ":" & PatRow).EntireRow.Delete
NotSaved:
    Unload PatFrm
    Patient_RefreshList  ' Corrected from Patients_RefreshList
End Sub
