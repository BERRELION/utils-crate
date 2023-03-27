Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim keywords As Variant
    Dim i As Long
    Dim keywordFile As String, keywordText As String, keywordList() As String
    
    ' Set the path to the file containing the keywords
    keywordFile = Environ("userprofile") & "\outlook_attachment_keywords.txt"
    
    ' If the file does not exist, create it with default contents
    If Not CreateObject("Scripting.FileSystemObject").FileExists(keywordFile) Then
        keywordText = "attach" & vbNewLine & "file" & vbNewLine & "document"
        CreateObject("Scripting.FileSystemObject").CreateTextFile(keywordFile).Write keywordText
    End If
    
    ' Read the keywords from the file into an array
    keywordText = CreateObject("Scripting.FileSystemObject").OpenTextFile(keywordFile, 1).ReadAll
    keywordList = Split(keywordText, vbNewLine)
    
    ' Loop through the keywords and check for attachments
    For i = LBound(keywordList) To UBound(keywordList)
        If InStr(1, Item.Body, keywordList(i), vbTextCompare) > 0 Or InStr(1, Item.Subject, keywordList(i), vbTextCompare) > 0 Then
            If Item.Attachments.Count = 0 Then
                answer = MsgBox("There's no attachment, send anyway?", vbYesNo)
                If answer = vbNo Then Cancel = True
                Exit For
            End If
        End If
    Next i
End Sub
