Attribute VB_Name = "BOXAuth"
Dim ie As New InternetExplorer
Function GetBoxAuthToken() As dictionary
Dim html As HTMLDocument

ie.Visible = True
ie.Navigate "https://account.box.com/api/oauth2/authorize?response_type=code&client_id=" & client_id & "&redirect_uri=" & redirect_uri & "&State=" & security_token
PageLoading
Set html = ie.document
Do While ie.LocationURL = "https://account.box.com/api/oauth2/authorize?response_type=code&client_id=" & client_id & "&redirect_uri=" & redirect_uri & "&State=" & security_token
DoEvents
Loop
PageLoading

Do While Left(ie.LocationURL, 43) = "https://app.box.com/api/oauth2/authorize?a="
DoEvents
Loop
PageLoading
GetUrl = ie.LocationURL
ie.Quit

Dim response$, sPostData$
Dim ohttp As Object
GetUrl = Mid(GetUrl, InStrRev(GetUrl, "code=") + 5)
Set ohttp = CreateObject("Microsoft.XMLHTTP")
    sPostData = "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""grant_type""" & vbCrLf & vbCrLf & _
        "authorization_code" & vbCrLf & _
        "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""code""" & vbCrLf & vbCrLf & _
        GetUrl & vbCrLf & _
        "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""client_id""" & vbCrLf & vbCrLf & _
        client_id & vbCrLf & _
        "--" & boundaryStr & vbCrLf & _
        "Content-Disposition: form-data; name=""client_secret""" & vbCrLf & vbCrLf & _
        client_secret & vbCrLf & _
        "--" & boundaryStr & "--"

ohttp.Open "POST", "https://api.box.com/oauth2/token", False
ohttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundaryStr
ohttp.send (sPostData)
response = ohttp.ResponseText
Debug.Print response
Debug.Print readAuth(response).Item("access_token")
Debug.Print readAuth(response).Item("refresh_token")
Set GetBoxAuthToken = readAuth(response)
Set GetBoxAuthToken = readAuth(response)
GetBoxAuthToken.Add "ConnectionStatusBOX", "ConnectedToBOX"
End Function

Sub PageLoading()
Do While ie.readyState <> READYSTATE_COMPLETE
DoEvents
Loop
End Sub

Function readAuth(ByVal strJSON As String) As dictionary
Set readAuth = New dictionary
'Creating collection keys: access_token, expires_in, restricted_to, refresh_token, token_type
i = 1
Do While InStr(1, strJSON, ",") > 0
    j = InStr(i, strJSON, ":")
    getleft = Mid(strJSON, i + 1, j - i - 1)
    If InStr(j, strJSON, ",") > 0 Then
        getright = Mid(strJSON, j + 1, InStr(j, strJSON, ",") - j - 1)
        readAuth.Add Replace(getleft, Chr(34), ""), Replace(getright, Chr(34), "")
    Else
        getright = Mid(strJSON, j + 1, InStr(j, strJSON, "}") - j - 1)
        readAuth.Add Replace(getleft, Chr(34), ""), Replace(getright, Chr(34), "")
        Exit Do
    End If
    strJSON = Mid(strJSON, InStr(j, strJSON, ","))
Loop
End Function

