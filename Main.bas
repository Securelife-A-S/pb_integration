Option Explicit

Sub DeleteVBComponent(ByVal CompName As String)

'Disabling the alert message
Application.DisplayAlerts = False

'Ignore errors
On Error Resume Next


'Delete the component
Dim vbCom As Object
     
Set vbCom = Application.VBE.ActiveVBProject.VBComponents

vbCom.Remove VBComponent:= _
vbCom.Item(CompName)
On Error GoTo 0

'Enabling the alert message
Application.DisplayAlerts = True

End Sub


Sub addBasFile(strPath As String)

Dim path As String
Dim objModule As Object

path = strPath & "\pb_integration-main\pensionBrokerExport.bas"
Call convertFile (path)
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "PensionBrokerExport"
Debug.Print ("PensionBrokerExport imported")

End Sub
Sub addUserForm(strPath As String)

Dim path As String
Dim objModule As Object
path = strPath & "\pb_integration-main\UserForm1.frm"
Call convertFile (path)


Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "UserForm1"
Debug.Print ("UserForm1 imported")
End Sub

Sub convertFile(path)
'Read content of the file with utf-8 encoding
Dim Content As String
With CreateObject("ADODB.Stream")
    .Type = 2  ' Private Const adTypeText = 2
    .Charset = "utf-8"
    .Open
    .LoadFromFile path
    Content = .ReadText(-1)  ' Private Const adReadAll = -1
    .Close
End With

'Replace Unix-style line endings with Windows-style line endings (Need to check if that applies to your file)
If InStr(Content, Chr$(13) & Chr$(10)) = 0 Then
    Content = Replace(Content, Chr$(10), Chr$(13) & Chr$(10))
End If

'Write file with default local ANSI encoding (generally Windows-1252 on Western/U.S. systems)
Open path For Output As #1
Print #1, Content
Close #1


End Sub

Sub init()
Dim strPath As String
strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\pb"

Call DeleteVBComponent("PensionBrokerExport")
Call DeleteVBComponent("UserForm1")

Call addBasFile(strPath)
Call addUserForm(strPath)
MsgBox "Arket er nu opdateret til den seneste version"

End Sub
