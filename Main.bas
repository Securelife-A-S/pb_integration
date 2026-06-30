Option Explicit

' === Selv-opdatering (pålidelig) =============================================
' Modulernes kode erstattes IN-PLACE i stedet for at blive fjernet og genimporteret.
'
' Hvorfor: VBComponents.Remove bliver først committet når makroen stopper med at køre.
' Den gamle løsning fjernede et modul og genimporterede det i SAMME kørsel, hvorefter
' "objModule.Name = ..." kolliderede med det endnu-ikke-fjernede modul. Resultatet var
' enten en runtime-fejl eller at den nye import fik et andet navn, mens DEN GAMLE KODE
' blev liggende. Det er årsagen til "opdateringen hentes, men der ligger stadig gammel kode".
'
' In-place erstatning (DeleteLines + AddFromString) er synkron og rammer ikke det problem.
' For UserForm røres kun kode-bagved - den binære .frx (layout/kontroller) bevares.
' =============================================================================

Sub init()
    Dim strPath As String, base As String, remoteVer As String
    strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\pb"
    base = strPath & "\pb_integration-main\"

    On Error GoTo UpdateFailed

    ' Erstat app-modulerne in-place (eller importér ved første installation).
    ReplaceComponentCode "PensionBrokerExport", base & "pensionBrokerExport.bas"
    ReplaceComponentCode "UserForm1", base & "UserForm1.frm"

    ' Retry-sikring: registrér først den installerede version når ALT er lykkedes.
    ' Fejler et trin ovenfor, sættes propertyen ikke, og opdateringen forsøges igen
    ' næste gang arket åbnes (i stedet for at brugeren sidder fast på gammel kode).
    remoteVer = ReadFileText(base & "version.txt")
    SetInstalledVersion remoteVer

    MsgBox "Arket er nu opdateret til den seneste version (" & Trim(remoteVer) & ")", vbInformation
    Exit Sub

UpdateFailed:
    MsgBox "Opdateringen fejlede (" & Err.Number & ": " & Err.Description & ")." & vbNewLine & _
           "Den nuværende version bevares, og opdateringen forsøges igen næste gang arket åbnes.", vbExclamation
End Sub

' Erstat en komponents kode in-place. Findes komponenten ikke (første installation),
' importeres filen i stedet - det medbringer .frx-layout for UserForms, og der er ingen
' gammel kode at kollidere med.
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

' Læs fil som UTF-8 og normalisér linjeendelser til CRLF (kræves af kodemodulet).
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

' Skær eksport-headeren væk så kun selve koden indsættes:
'   .frm/.cls -> VERSION-linje + Begin..End design-blok + Attribute-linjer fjernes.
'   .bas (uden header) -> hele indholdet returneres.
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
                ' inde i design-blok, skip
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

' Installeret-version gemt som workbook-property (grundlag for retry-sikringen).
' CustomDocumentProperties er typet "As Object" i Excel-biblioteket, så det virker uden
' en eksplicit reference til Office-biblioteket; Type 4 = msoPropertyTypeString.
Function InstalledVersion() As String
    Dim v As String
    v = ""
    On Error Resume Next
    v = ThisWorkbook.CustomDocumentProperties("pb_installed_version").value
    On Error GoTo 0
    InstalledVersion = v
End Function

Sub SetInstalledVersion(ver As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("pb_installed_version").value = Trim(ver)
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add _
            Name:="pb_installed_version", LinkToContent:=False, _
            Type:=4, value:=Trim(ver)
    End If
    On Error GoTo 0
End Sub
