Option Explicit

' === Bootstrap for selv-opdatering ===========================================
' Denne kode ligger i ThisWorkbook og kan ikke auto-opdatere sig selv. Den holdes
' derfor tynd: tjek version -> download -> erstat "Main" in-place -> kør Main.init,
' som klarer resten (PensionBrokerExport + UserForm1) på samme pålidelige måde.
'
' Bemærk: eksisterende ark i marken har en ÆLDRE udgave af denne fil og kan ikke få
' denne forbedring via auto-update. De henter dog stadig den nye Main.bas og kalder
' Main.init, så den pålidelige in-place erstatning af app-modulerne når ud til dem.
' =============================================================================

Function getRemoteVersion() As String
    Dim http As Object
    getRemoteVersion = ""
    On Error GoTo fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://raw.githubusercontent.com/Securelife-A-S/pb_integration/main/version.txt", False
    http.send
    If http.Status = 200 Then getRemoteVersion = http.responseText
fail:
End Function

' Installeret-version som workbook-property (sættes af Main.init når opdateringen er
' fuldt gennemført). Mangler den (første kørsel), returneres "".
Function InstalledVersion() As String
    Dim v As String
    v = ""
    On Error Resume Next
    v = ThisWorkbook.CustomDocumentProperties("pb_installed_version").value
    On Error GoTo 0
    InstalledVersion = v
End Function

' Sandt hvis Excel har "Hav tillid til adgang til VBA-projektobjektmodellen" slået til.
Function VBProjectAccessible() As Boolean
    Dim n As Long
    VBProjectAccessible = False
    On Error GoTo done
    n = ThisWorkbook.VBProject.VBComponents.Count
    VBProjectAccessible = True
done:
End Function

Function VbomHelpText() As String
    VbomHelpText = _
        "Opdateringen kunne ikke gennemføres, fordi Excel ikke har tillid til adgang til VBA-projektet." & vbNewLine & vbNewLine & _
        "Slå indstillingen til:" & vbNewLine & _
        "Filer  >  Indstillinger  >  Sikkerhedscenter  >  Indstillinger for Sikkerhedscenter  >  Indstillinger for makroer" & vbNewLine & _
        "-> sæt flueben i 'Hav tillid til adgang til VBA-projektobjektmodellen'." & vbNewLine & vbNewLine & _
        "Genstart derefter Excel og åbn arket igen."
End Function

Function MkDir(strPath As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(strPath) Then
        FSO.DeleteFolder strPath, True
        Debug.Print "Deleting folder"
    End If
    If Not FSO.FolderExists(strPath) Then
        FSO.CreateFolder strPath
        Debug.Print "Creating folder"
    End If
End Function

Function downloadAndUnzip(strPath As String)
    Dim FileUrl As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object

    FileUrl = "https://github.com/Securelife-A-S/pb_integration/archive/refs/heads/main.zip"
    Set objXmlHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objXmlHttpReq.Open "GET", FileUrl, False
    objXmlHttpReq.send

    If objXmlHttpReq.Status <> 200 Then
        Err.Raise vbObjectError + 513, "downloadAndUnzip", "Kunne ikke hente opdateringen (HTTP " & objXmlHttpReq.Status & ")."
    End If

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objXmlHttpReq.responseBody
    objStream.SaveToFile strPath & "\pb.zip", 2
    objStream.Close
    Debug.Print ("Download done")

    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(strPath & "\").CopyHere ShellApp.Namespace(strPath & "\pb.zip").Items

    ' CopyHere er asynkron - vent til den udpakkede version.txt faktisk findes,
    ' ellers læser erstatningen bagefter filer der ikke er der endnu.
    Dim FSOWait As Object
    Dim extractedFile As String
    Dim waitCount As Integer
    Set FSOWait = CreateObject("Scripting.FileSystemObject")
    extractedFile = strPath & "\pb_integration-main\version.txt"
    waitCount = 0
    Do Until FSOWait.FileExists(extractedFile) Or waitCount >= 60
        Application.Wait Now + TimeValue("0:00:01")
        DoEvents
        waitCount = waitCount + 1
    Loop

    If Not FSOWait.FileExists(extractedFile) Then
        Err.Raise vbObjectError + 514, "downloadAndUnzip", "Udpakning af opdateringen tog for lang tid eller fejlede."
    End If
    Debug.Print ("Unpack done")
End Function

' Erstat en komponents kode in-place (samme princip som i Main.bas). Bruges her kun til
' at installere/erstatte "Main". Findes komponenten ikke, importeres filen.
Sub ReplaceComponentCode(compName As String, filePath As String)
    Dim comp As Object, body As String

    Set comp = Nothing
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(compName)
    On Error GoTo 0

    If comp Is Nothing Then
        Set comp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
        comp.Name = compName
    Else
        body = ExtractCodeBody(ReadFileText(filePath))
        With comp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString body
        End With
    End If
End Sub

Function ReadFileText(path As String) As String
    Dim s As String
    With CreateObject("ADODB.Stream")
        .Type = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile path
        s = .ReadText(-1)  ' adReadAll
        .Close
    End With
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    ReadFileText = s
End Function

Function ExtractCodeBody(content As String) As String
    Dim lines() As String, i As Long, depth As Long
    Dim started As Boolean, t As String, out As String
    lines = Split(content, vbCrLf)
    started = False
    depth = 0
    For i = LBound(lines) To UBound(lines)
        If started Then
            out = out & lines(i) & vbCrLf
        Else
            t = Trim(lines(i))
            If Left(t, 8) = "VERSION " Then
                ' design-header, skip
            ElseIf Left(t, 6) = "Begin " Or t = "Begin" Then
                depth = depth + 1
            ElseIf depth > 0 Then
                If t = "End" Then depth = depth - 1
            ElseIf Left(t, 10) = "Attribute " Then
                ' modul-attribut, skip
            ElseIf t = "" Then
                ' ledende tomme linjer, skip
            Else
                started = True
                out = out & lines(i) & vbCrLf
            End If
        End If
    Next i
    ExtractCodeBody = out
End Function

Sub Workbook_Open()
    Dim strPath As String, base As String, remoteVer As String
    strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\pb"
    base = strPath & "\pb_integration-main\"

    remoteVer = getRemoteVersion()
    If Trim(remoteVer) = "" Then Exit Sub                                     ' offline - prøv igen senere
    If StrComp(Trim(remoteVer), Trim(InstalledVersion())) = 0 Then Exit Sub   ' allerede opdateret

    ' Tjek VBA-adgang FØR download, så disk-versionen ikke avancerer ved fejl
    ' (ellers ville opdateringen tro den var færdig og aldrig forsøge igen).
    If Not VBProjectAccessible() Then
        MsgBox VbomHelpText(), vbExclamation, "Opdatering kræver VBA-adgang"
        Exit Sub
    End If

    On Error GoTo UpdateFailed
    MsgBox "Der er kommet ny version - opdatering påbegyndt", vbInformation
    Call MkDir(strPath)
    Call downloadAndUnzip(strPath)
    ' Installér/erstat opdateringslogikken (Main) in-place og lad Main.init klare resten.
    Call ReplaceComponentCode("Main", base & "Main.bas")
    Application.Run "Main.init"
    Exit Sub

UpdateFailed:
    MsgBox "Opdateringen kunne ikke fuldføres (" & Err.Number & ": " & Err.Description & ")." & vbNewLine & _
           "Forsøges igen næste gang arket åbnes.", vbExclamation
End Sub
