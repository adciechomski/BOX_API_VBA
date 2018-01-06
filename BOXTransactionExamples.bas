Attribute VB_Name = "BOXTransactionExamples"
Option Explicit
'###########################Functionality summary############################################
'
'This module presents utilization of BOXAuth module and BOXFileUpload module
'and how it can be used to interact with BOX doing HTTP calls using VBA.
'BOXAuth module is crucial to do all later calls, whereas BOXFileUpload module is example of POST API call
'
'Please, remmember to add appropierte references to your project: ScriptingRuntime, Microsoft HTTP Object Library, Microsoft Internet Controls
'
'###########################################################################################
'##############here are example of ContentType notation for sample file types, found more on web ##################
'"application/vnd.ms-excel.sheet.binary.macroenabled.12" - binary excel file contentType
'"plain/text"- text file contentType
'###################################### argumentsfor declared here for testing purposes #####################
Global Const boundaryStr  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
Global Const client_id  As String = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
Global Const client_secret  As String = "yyyyyyyyyyyyyyyyyyyyyyyyyyyyyy"
Global Const security_token As String = "security_token%3DKnhMJatFipTtyuiopqwdc"
Global Const redirect_uri As String = "http://0.0.0.0"
'###################################### / argumentsfor declared here for testing purposes #####################
Dim accessTokenDict As New dictionary

Sub UploadingFileSample()
Dim response As String
Dim Key As Variant

Set accessTokenDict = GetBoxAuthToken ' Getting access Token needed for BOX transactions

For Each Key In accessTokenDict
    Debug.Print Key & ":" & accessTokenDict.Item(Key)
Next Key

response = pvPostFileBinaryXLS("https://upload.box.com/api/2.0/files/content" _
    , "C:\binaryFile.xlsb" _
    , "application/vnd.ms-excel.sheet.binary.macroenabled.12" _
    , accessTokenDict.Item("access_token") _
    , "0" _
    , False)

Debug.Print response
End Sub
