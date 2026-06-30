VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
   Caption         =   "UserForm1"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7520
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === Importér medarbejder fra admin-portal-v2 ================================
' Login med portal-email + adgangskode (POST /api/excel/login) -> token i
' session-hukommelse -> opslag (GET /api/excel/employee/{cpr}) med Bearer-token.
' apikeyBox (manuelt token) er udfaset. Email/password-felterne tilføjes ved
' runtime (Controls.Add), så der ikke kræves .frx-ændringer i Excel.
' ===========================================================================

Private mToken As String          ' cachet token (session)
Private mEmail As String          ' husket email (session) - password gemmes ALDRIG
Private mEmailBox As MSForms.TextBox
Private mPwdBox As MSForms.TextBox
Private mStatus As MSForms.Label

Private Sub UserForm_Initialize()
    On Error GoTo buildErr
    Dim ctl As MSForms.Control

    ' Ryd gamle labels (vi tegner vores egne) + skjul udfasede kontroller.
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Label" Then ctl.Visible = False
    Next ctl
    apikeyBox.Visible = False
    ExportButton.Visible = False

    Me.Caption = "Importér medarbejder fra portalen"
    Me.Width = 252
    Me.Height = 312

    AddLabel "lblMiljo", "Vælg miljø:", 12, 6, 228
    PlaceRadio BedstPensionButton, "CPOF", 18, 22
    PlaceRadio SecureLifeTestButton, "CPOF test", 18, 40
    PlaceRadio SecureLifeButton, "SecureLife (ikke klar)", 18, 58
    SecureLifeButton.Enabled = False
    BedstPensionButton.value = True            ' fornuftig standard

    AddLabel "lblCpr", "CPR-nummer (10 cifre, uden bindestreg):", 12, 82, 228
    PlaceBox cprBox, 12, 96

    AddLabel "lblEmail", "Portal-email:", 12, 122, 228
    If mEmailBox Is Nothing Then Set mEmailBox = Me.Controls.Add("Forms.TextBox.1", "emailBox", True)
    PlaceBox mEmailBox, 12, 136
    mEmailBox.value = mEmail

    AddLabel "lblPwd", "Adgangskode:", 12, 162, 228
    If mPwdBox Is Nothing Then Set mPwdBox = Me.Controls.Add("Forms.TextBox.1", "passwordBox", True)
    PlaceBox mPwdBox, 12, 176
    mPwdBox.PasswordChar = "*"

    If mStatus Is Nothing Then Set mStatus = Me.Controls.Add("Forms.Label.1", "statusLabel", True)
    mStatus.Left = 12: mStatus.Top = 202: mStatus.Width = 228: mStatus.Height = 26
    mStatus.Caption = ""

    PlaceButton ImportButton, "Hent medarbejder", 12, 234
    ImportButton.Default = True
    Exit Sub

buildErr:
    MsgBox "Kunne ikke opbygge login-formularen: " & Err.Description, vbCritical, "Fejl"
End Sub

' --- Klik: validér -> (login) -> hent -> skriv til ark ----------------------
Private Sub ImportButton_Click()
    Dim baseUrl As String, cpr As String, email As String, password As String
    Dim token As String, emp As Object
    Dim code As Long, body As String, attempt As Integer

    SetStatus ""

    If Not EnvSelected(baseUrl) Then Warn "Vælg et miljø (CPOF eller CPOF test).": Exit Sub

    cpr = CleanCpr(cprBox.value)
    If Len(cpr) <> 10 Then Warn "CPR skal være 10 cifre (uden bindestreg).": cprBox.SetFocus: Exit Sub

    email = Trim(mEmailBox.value)
    password = mPwdBox.value
    If InStr(email, "@") = 0 Then Warn "Indtast din portal-email.": mEmailBox.SetFocus: Exit Sub
    If Len(password) = 0 Then Warn "Indtast din adgangskode.": mPwdBox.SetFocus: Exit Sub

    ImportButton.Enabled = False

    ' Token: brug cachet, ellers log ind.
    token = mToken
    If Len(token) = 0 Then
        SetStatus "Logger ind ..."
        token = ExcelLogin(baseUrl, email, password)
        If Len(token) = 0 Then GoTo done       ' ExcelLogin har vist fejlen
    End If

    ' Hent medarbejder, med ét gen-login-forsøg ved udløbet token (401).
    For attempt = 1 To 2
        SetStatus "Henter medarbejder ..."
        Set emp = HttpGetEmployee(baseUrl, cpr, token, code, body)
        If code = 200 Then Exit For
        If code = 401 And attempt = 1 Then
            mToken = ""
            SetStatus "Sessionen udløb - logger ind igen ..."
            token = ExcelLogin(baseUrl, email, password)
            If Len(token) = 0 Then GoTo done
        Else
            Warn EmployeeErr(code, body): GoTo done
        End If
    Next attempt

    If emp Is Nothing Then Warn "Uventet tomt svar fra portalen.": GoTo done

    WriteEmployee emp

done:
    ImportButton.Enabled = True
End Sub

' --- Miljø-valg -> base-URL -------------------------------------------------
Private Function EnvSelected(ByRef baseUrl As String) As Boolean
    EnvSelected = True
    If SecureLifeTestButton.value = True Then
        baseUrl = "https://test.portal.cpof.dk"    ' CPOF test
    ElseIf BedstPensionButton.value = True Then
        baseUrl = "https://portal.cpof.dk"         ' CPOF (prod)
    ElseIf SecureLifeButton.value = True Then
        baseUrl = "https://portal.securelife.dk"   ' SecureLife (ikke klar)
    Else
        EnvSelected = False
    End If
End Function

' --- Login: bytter email+password til et token -----------------------------
Private Function ExcelLogin(baseUrl As String, email As String, password As String) As String
    Dim http As Object, code As Long, body As String, j As Object
    ExcelLogin = ""
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo neterr
    http.Open "POST", baseUrl & "/api/excel/login", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.send "{""email"":""" & JsonEsc(email) & """,""password"":""" & JsonEsc(password) & """}"
    On Error GoTo 0

    code = http.Status
    body = http.responseText
    Select Case code
        Case 200
            Set j = JsonConverter.ParseJson(body)
            ExcelLogin = CStr(j("token"))
            mToken = ExcelLogin
            mEmail = email
        Case 401: Warn "Forkert email eller adgangskode."
        Case 403: Warn "Din bruger har ikke adgang til Excel-integrationen."
        Case 429: Warn "For mange loginforsøg. Vent lidt og prøv igen."
        Case Else: Warn JsonMessageOr(body, "Login fejlede (HTTP " & code & ").")
    End Select
    Exit Function

neterr:
    Warn "Kunne ikke kontakte portalen. Tjek din internetforbindelse."
End Function

' --- Opslag: henter medarbejderen (returnerer Nothing ved fejl) -------------
Private Function HttpGetEmployee(baseUrl As String, cpr As String, token As String, _
                                 ByRef code As Long, ByRef body As String) As Object
    Dim http As Object
    Set HttpGetEmployee = Nothing
    code = 0: body = ""
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo neterr
    http.Open "GET", baseUrl & "/api/excel/employee/" & cpr, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & token
    http.send
    On Error GoTo 0

    code = http.Status
    body = http.responseText
    If code = 200 Then Set HttpGetEmployee = JsonConverter.ParseJson(body)
    Exit Function

neterr:
    code = -1
End Function

Private Function EmployeeErr(code As Long, body As String) As String
    Select Case code
        Case -1: EmployeeErr = "Kunne ikke kontakte portalen. Tjek din internetforbindelse."
        Case 401: EmployeeErr = "Login udløb. Prøv igen."
        Case 403: EmployeeErr = "Du har ikke adgang til denne medarbejders virksomhed."
        Case 404: EmployeeErr = JsonMessageOr(body, "Medarbejderen blev ikke fundet.")
        Case Else: EmployeeErr = JsonMessageOr(body, "Opslaget fejlede (HTTP " & code & ").")
    End Select
End Function

' --- Skriv til arket "Portal import-eksport" (uændret nøgle-mapping) --------
Private Sub WriteEmployee(emp As Object)
    Dim sheetName As String, ws As Worksheet
    Dim preview As String, k As Variant, v As Variant
    Dim ans As Integer, i As Long, keyName As String

    sheetName = "Portal import-eksport"
    If Not sheetExist(sheetName) Then
        Warn "Arket '" & sheetName & "' findes ikke i denne projektmappe."
        Exit Sub
    End If
    Set ws = ActiveWorkbook.Sheets(sheetName)

    For Each k In emp.keys
        If IsNull(emp(k)) Then v = "" Else v = CStr(emp(k))
        preview = preview & k & ": " & v & vbNewLine
    Next k

    ans = MsgBox(preview, vbQuestion + vbYesNo + vbDefaultButton1, "Importér denne medarbejder?")
    If ans <> vbYes Then SetStatus "Import annulleret.": Exit Sub

    For i = 2 To 60
        keyName = CStr(ws.Cells(i, 1).value)
        If Len(keyName) > 0 Then
            v = ""
            If emp.Exists(keyName) Then
                If Not IsNull(emp(keyName)) Then v = emp(keyName)
            End If
            ws.Cells(i, 2).value = v
        End If
    Next i

    SetStatus "Import OK."
    MsgBox "Medarbejderen er importeret til arket '" & sheetName & "'.", vbInformation, "Import OK"
    Me.Hide
End Sub

Private Sub ExportButton_Click()
    ' admin-portal-v2-integrationen er skrivebeskyttet (intet PUT-endpoint).
    MsgBox "Eksport til portalen understøttes ikke i denne version (skrivebeskyttet integration).", vbInformation
End Sub

' --- Hjælpere ---------------------------------------------------------------
Function sheetExist(sSheet As String) As Boolean
    On Error GoTo ErrorMSG
    sheetExist = (ActiveWorkbook.Sheets(sSheet).index > 0)
    Exit Function
ErrorMSG:
    sheetExist = False
End Function

Private Function CleanCpr(s As String) As String
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then CleanCpr = CleanCpr & c
    Next i
End Function

Private Function JsonEsc(s As String) As String
    Dim r As String
    r = s
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCr, "\r")
    r = Replace(r, vbLf, "\n")
    r = Replace(r, vbTab, "\t")
    JsonEsc = r
End Function

Private Function JsonMessageOr(body As String, fallback As String) As String
    Dim j As Object
    On Error GoTo fb
    If Len(Trim(body)) > 0 Then
        Set j = JsonConverter.ParseJson(body)
        If j.Exists("message") Then
            If Len(CStr(j("message"))) > 0 Then
                JsonMessageOr = CStr(j("message"))
                Exit Function
            End If
        End If
    End If
fb:
    JsonMessageOr = fallback
End Function

Private Sub SetStatus(msg As String)
    On Error Resume Next
    mStatus.ForeColor = RGB(0, 0, 0)
    mStatus.Caption = msg
    Me.Repaint
End Sub

Private Sub Warn(msg As String)
    On Error Resume Next
    mStatus.ForeColor = RGB(170, 0, 0)
    mStatus.Caption = msg
    On Error GoTo 0
    MsgBox msg, vbExclamation, "Importér medarbejder"
End Sub

Private Sub AddLabel(nm As String, cap As String, x As Single, y As Single, w As Single)
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1", nm, True)
    l.Caption = cap
    l.Left = x: l.Top = y: l.Width = w: l.Height = 12
End Sub

Private Sub PlaceBox(b As MSForms.Control, x As Single, y As Single)
    b.Left = x: b.Top = y: b.Width = 222: b.Height = 18
    b.Visible = True
End Sub

Private Sub PlaceRadio(r As MSForms.Control, cap As String, x As Single, y As Single)
    r.Caption = cap
    r.Left = x: r.Top = y: r.Width = 210: r.Height = 16
    r.Visible = True
End Sub

Private Sub PlaceButton(btn As MSForms.Control, cap As String, x As Single, y As Single)
    btn.Caption = cap
    btn.Left = x: btn.Top = y: btn.Width = 222: btn.Height = 26
    btn.Visible = True
End Sub
