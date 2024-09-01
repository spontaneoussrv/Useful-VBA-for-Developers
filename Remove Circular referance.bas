Attribute VB_Name = "Module4"
Sub RemoveCircularReferences()
    Dim ws As Worksheet
    Dim cell As Range
    Dim circularRefs As Range
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each cell in the sheet to check for circular references
        For Each cell In ws.UsedRange
            If Not IsError(cell.value) Then
                On Error Resume Next
                If cell.CircularReference.Address <> "" Then
                    If circularRefs Is Nothing Then
                        Set circularRefs = cell
                    Else
                        Set circularRefs = Union(circularRefs, cell)
                    End If
                End If
                On Error GoTo 0
            End If
        Next cell
    Next ws
    
    ' Clear the contents of cells with circular references
    If Not circularRefs Is Nothing Then
        circularRefs.ClearContents
        MsgBox "Circular references found and cleared."
    Else
        MsgBox "No circular references found."
    End If
End Sub

