Attribute VB_Name = "BOXFileUpload"
Public Function pvPostFileBinaryXLS(URL As String, filePath As String, contentType As String, authorizationToken As String, Optional folderID As String = "0", Optional ByVal bAsync As Boolean) As String
    Dim nFile As Integer
    Dim baBuffer() As Byte
    Dim sPostData$
        
    filePath = Mid$(filePath, InStrRev(filePath, "\") + 1)
    atrr = "{""name"":""" & filePath & """, ""parent"": {""id"": """ & folderID & """}}"
    '--- write file into Unicode using the default code page of the system
    nFile = FreeFile
    Open filePath For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    Close nFile
    '--- prepare body
    sPostData = "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""attributes""" & vbCrLf & vbCrLf & _
        atrr & vbCrLf & _
        "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & filePath & """" & vbCrLf & _
        "Content-Type: " & contentType & vbCrLf & vbCrLf & _
        sPostData & vbCrLf & _
        "--" & boundaryStr & "--"
    '--- post
    With CreateObject("Microsoft.XMLHTTP")
        .Open "POST", URL, bAsync
        .setRequestHeader "Authorization", "Bearer " & authorizationToken
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundaryStr
        .send pvToByteArray(sPostData)
        If Not bAsync Then
            pvPostFileBinaryXLS = .ResponseText
        End If
    End With
End Function
Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function


