Attribute VB_Name = "Disable_Cross_Button"
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    
    MsgBox "Please use the command button to close the form."
End Sub
