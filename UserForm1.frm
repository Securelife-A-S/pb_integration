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


Function sheetExist(sSheet As String) As Boolean
On Error GoTo ErrorMSG
sheetExist = (ActiveWorkbook.Sheets(sSheet).index > 0)
Exit Function

ErrorMSG:
MsgBox "Navnet på pensionsleverandøren findes ikke i arket" & sSheet & " - Har du husket at sætte den korrekt navn for leverandøren under dets 'EXCEL' felt?"

End Function



Private Sub ExportButton_Click()

' admin-portal-v2-integrationen er skrivebeskyttet (intet PUT-endpoint).
' Eksport til portalen understøttes ikke i denne version.
MsgBox "Eksport til portalen understøttes ikke i denne version (skrivebeskyttet integration).", vbInformation

End Sub




Private Sub ImportButton_Click()

Dim JsonObject As Object
Dim objRequest As Object
Dim strBaseUrl As String
Dim strUrl As String
Dim strResponse As String
Dim responseStatus As Variant

' Mandatory fields
If cprBox.value = "" Then
    MsgBox "Du skal indsætte cprnummer", vbCritical
    Exit Sub
End If

' apikeyBox bruges nu til portal-tokenet (genereres i portalen under /settings/api-tokens).
If apikeyBox.value = "" Then
    MsgBox "Du skal indsætte portal-token", vbCritical
    Exit Sub
End If

If SecureLifeButton.value = False And BedstPensionButton.value = False And SecureLifeTestButton.value = False Then
    MsgBox "Du skal vælge miljø", vbCritical
    Exit Sub
End If

If BedstPensionButton.value = True Then
    MsgBox "BedstPension er endnu ikke migreret til den nye portal.", vbInformation
    Exit Sub
End If

' Map miljø-valg til admin-portal-v2 base-URL.
If SecureLifeTestButton.value = True Then
    strBaseUrl = "https://test2.portal.cpof.dk" ' TODO: bekræft public staging-URL
Else
    strBaseUrl = "https://portal.cpof.dk" ' CPOF (prod) — aktiv bruger
End If

strUrl = strBaseUrl & "/api/excel/employee/" & cprBox.value

Set objRequest = CreateObject("MSXML2.XMLHTTP")
With objRequest
    .Open "GET", strUrl, False
    .setRequestHeader "Accept", "application/json"
    .setRequestHeader "Authorization", "Bearer " & apikeyBox.value
    .send
    strResponse = .responseText
    responseStatus = .Status
End With
Debug.Print (responseStatus)

If Not responseStatus = 200 Then
    ' 401 = token udløbet/ugyldigt -> generér nyt token i portalen og indsæt igen.
    MsgBox "ERROR: " & strResponse
    Exit Sub
Else
    Set JsonObject = JsonConverter.ParseJson(strResponse)
End If


Dim text As String
' Generate a text of key value pair for popup & printing the dict
For Each key In JsonObject.keys
    Dim val As Variant
    If IsNull(JsonObject(key)) Then
        val = ""
    Else
        val = CStr(JsonObject(key))
    End If
    text = text + " " + key + ": " + val + vbNewLine
Next

Dim sheetName As String
sheetName = "Portal import-eksport"
Dim worksheet As worksheet

Set worksheet = Sheets(sheetName)

Dim answer As Integer
answer = MsgBox(text, vbQuestion + vbYesNo + vbDefaultButton2, "Import af medarbejderen")
If answer = vbYes Then
    Dim cellIndex As Integer
    cellIndex = 2
    ' Adding data to Stamoplysninger sheet
    For i = 2 To 60
        worksheet.Cells(cellIndex, 2).value = JsonObject(worksheet.Cells(cellIndex, 1).value)
        cellIndex = cellIndex + 1
    Next i
    
    MsgBox "Import OK"
    UserForm1.Hide
Else
  MsgBox "Import fejlede"
End If


End Sub



