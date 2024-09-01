Attribute VB_Name = "HighlightCircularReference"
Sub HighlightCircularReferences()
    Dim ws As Worksheet
    Dim cell As Range
    Dim circularRefs As Range
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each cell in the sheet to check for circular references
        For Each cell In ws.UsedRange
            If Not IsError(cell.Value) Then
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
    
    ' Highlight cells with circular references in yellow
    If Not circularRefs Is Nothing Then
        circularRefs.Interior.Color = RGB(255, 255, 0) ' Yellow color
        MsgBox "Circular references found and highlighted."
    Else
        MsgBox "No circular references found."
    End If
End Sub

