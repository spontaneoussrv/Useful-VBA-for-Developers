Attribute VB_Name = "Protect_and_Hide_Sheets"
Sub ProtectAndHideSheet()
 Application.ScreenUpdating = False
 
    Dim ws As Worksheet
    ' Specify the name of the sheet you want to protect and hide
        Set ws = ThisWorkbook.Sheets("Sheet2")
    ' protect the sheet if it's protected
        ws.Protect Password:="8246", UserInterfaceOnly:=True
    ' Hide the sheet
        ws.Visible = xlSheetVeryHidden

    

 Application.ScreenUpdating = True
End Sub


Sub UnProtectAndUnHideSheet()
 Application.ScreenUpdating = False

    Dim ws As Worksheet
    ' Specify the name of the sheet you want to protect and hide
        Set ws = ThisWorkbook.Sheets("Sheet2")
    ' UnProtect the sheet
        If ws.ProtectContents Then ws.Unprotect Password:="8246"
    ' Unhide the sheet
        If ws.Visible = xlSheetVeryHidden Then ws.Visible = xlSheetVisible
        
        
    
 Application.ScreenUpdating = True
End Sub


