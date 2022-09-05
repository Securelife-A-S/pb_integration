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


Function addBasFile()

Dim path As String
Dim objModule As Object

path = ThisWorkbook.path & "\pb\pb_integration-main\pensionBrokerExport.bas"
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "PensionBrokerExport"


Debug.Print path

End Function
Function addUserForm()

Dim path As String
Dim objModule As Object

path = ThisWorkbook.path & "\pb\pb_integration-main\UserForm1.frm"
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "UserForm1"

Debug.Print path

End Function



Private Sub init()

Call DeleteVBComponent("PensionBrokerExport")
Call DeleteVBComponent("UserForm1")
Call addBasFile
Call addUserForm
MsgBox "Rådgivningsværktøjet er nu opdateret til den seneste version"


End Sub
