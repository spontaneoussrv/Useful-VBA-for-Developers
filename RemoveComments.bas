Attribute VB_Name = "RemoveComments"
'activate from referance Microsoft Visual Basic for Applications Extensibility 5.3.

Sub RemoveAllComments()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim vbMod As Object
    Dim i As Long
    Dim lineText As String
    Dim lineCount As Long
    
    
    Set vbProj = ThisWorkbook.VBProject

    ' Check if the reference is already added
    Dim ref As Object
    Dim refFound As Boolean
    refFound = False

    For Each ref In vbProj.References
        If ref.Name = "VBIDE" Then
            refFound = True
            Exit For
        End If
    Next ref

    ' Add the reference if not already added
    If Not refFound Then
        vbProj.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
        MsgBox "Microsoft Visual Basic for Applications Extensibility 5.3 reference added."
    Else
      '  MsgBox "Reference is already added."
    End If

    Set vbProj = ActiveWorkbook.VBProject

    For Each vbComp In vbProj.VBComponents
        Set vbMod = vbComp.CodeModule

        lineCount = vbMod.CountOfLines

        For i = lineCount To 1 Step -1
            lineText = vbMod.lines(i, 1)
            
            If Trim(lineText) Like "'*" Then
                vbMod.DeleteLines i
            End If
        Next i
    Next vbComp
    
    MsgBox "All comments removed from the VBA project.", vbInformation
End Sub
