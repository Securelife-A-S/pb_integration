Option Explicit

Function DeleteVBComponent(ByVal CompName As String)

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

End Function


Function addBasFile(strPath As String)

Dim path As String
Dim objModule As Object

path = strPath & "\pb_integration-main\pensionBrokerExport.bas"
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "PensionBrokerExport"
Debug.Print ("PensionBrokerExport imported")

End Function
Function addUserForm(strPath As String)

Dim path As String
Dim objModule As Object

path = strPath & "\pb_integration-main\UserForm1.frm"
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "UserForm1"
Debug.Print ("UserForm1 imported")
End Function



Sub init()
Dim strPath As String
strPath = Environ("USERPROFILE") & "\Desktop\pb"
Call DeleteVBComponent("PensionBrokerExport")
Call DeleteVBComponent("UserForm1")
Call addBasFile(strPath)
Call addUserForm(strPath)
MsgBox "Arket er nu opdateret til den seneste version"


End Sub
