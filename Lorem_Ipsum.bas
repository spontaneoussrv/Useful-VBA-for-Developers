Attribute VB_Name = "Lorem_Ipsum"

Function GenerateLoremIpsum(WordCount As Integer) As String
    Dim LoremIpsumText As String
    LoremIpsumText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " & _
                     "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. " & _
                     "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. " & _
                     "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
    
    Dim Words() As String
    Words = Split(LoremIpsumText, " ")
    
    Dim i As Integer
    Dim Result As String
    Result = ""
    
    For i = 0 To WordCount - 1
        Result = Result & Words(i Mod UBound(Words) + 1) & " "
    Next i
    
    GenerateLoremIpsum = Trim(Result)
End Function


