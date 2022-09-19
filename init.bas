
Option Explicit

Function versionIsOutdated(strPath As String)

Dim FSO As New FileSystemObject
Dim FileToRead As Variant
Dim TextString As String

If FSO.FolderExists(strPath) Then
' exist, lookup versionNumber
        Dim FileUrl As String
        Dim objXmlHttpReq As Object
        Dim objStream As Object
        Dim strResult
        
        FileUrl = "https://raw.githubusercontent.com/Securelife-A-S/pb_integration/main/version.txt"
        
        Set objXmlHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objXmlHttpReq.Open "GET", FileUrl, False
        objXmlHttpReq.send
        strResult = objXmlHttpReq.responseText
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set FileToRead = FSO.OpenTextFile(strPath & "\pb_integration-main\version.txt", ForReading) 'add here the path of your text file
        TextString = FileToRead.ReadAll
        FileToRead.Close
        Debug.Print (TextString)
        Debug.Print (strResult)
        Dim compResult As Integer
        
        If StrComp(TextString, strResult) = 0 Then
        Debug.Print ("Version is up to date")
        versionIsOutdated = False
        Else
        versionIsOutdated = True
        Debug.Print ("Version is outdated")
        End If
Else
versionIsOutdated = True
Debug.Print ("Folder is not downloaded yet")
End If

End Function
Function MkDir(strPath As String)

Dim FSO As New FileSystemObject

If FSO.FolderExists(strPath) Then
' exist, so delete the folder
          FSO.DeleteFolder strPath, True
          Debug.Print "Deleting folder"
End If

If Not FSO.FolderExists(strPath) Then

' doesn't exist, so create the folder
          FSO.CreateFolder strPath
          Debug.Print "Creating folder"
End If

End Function


Function downloadAndUnzip(strPath As String)


Dim FileUrl As String
Dim objXmlHttpReq As Object
Dim objStream As Object
Dim strResult

FileUrl = "https://raw.githubusercontent.com/Securelife-A-S/pb_integration/main/version.txt"

Set objXmlHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXmlHttpReq.Open "GET", FileUrl, False
objXmlHttpReq.send
strResult = objXmlHttpReq.responseText

Debug.Print (strResult)

FileUrl = "https://github.com/Securelife-A-S/pb_integration/archive/refs/heads/main.zip"

Set objXmlHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXmlHttpReq.Open "GET", FileUrl, False
objXmlHttpReq.send

If objXmlHttpReq.Status = 200 Then
     Set objStream = CreateObject("ADODB.Stream")
     objStream.Open
     objStream.Type = 1
     objStream.Write objXmlHttpReq.responseBody
     objStream.SaveToFile strPath & "\pb.zip", 2
     objStream.Close
End If


Debug.Print ("Download done")
     
Dim ShellApp As Object
'Copy the files & folders from the zip into a folder
Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(strPath & "\").CopyHere ShellApp.Namespace(strPath & "\pb.zip").Items

Debug.Print ("Unpack done")
End Function

Function DeleteVBComponent()
Dim CompName As String
CompName = "Main"
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

path = strPath & "\pb_integration-main\Main.bas"
Set objModule = Application.VBE.ActiveVBProject.VBComponents.Import(path)
objModule.Name = "Main"


Debug.Print path

End Function

Sub Workbook_Open()

Dim strPath As String
strPath = Environ("USERPROFILE") & "\Desktop\pb"
Debug.Print (strPath)
If versionIsOutdated(strPath) = True Then
    MsgBox "Der er kommet ny version - Downloading påbegyndt"
    Call MkDir(strPath)
    Call downloadAndUnzip(strPath)
    Call DeleteVBComponent
    Call addBasFile(strPath)
    Application.Run ("Main.init")
End If

End Sub



