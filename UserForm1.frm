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

Dim dict As New Scripting.Dictionary

Dim JsonObject As Object
Dim objRequest As Object
Dim strUrl As String
Dim blnAsync As Boolean
Dim strResponse As String
Dim codeResponse As String
Dim responseStatus As Variant


' Mandatory field
If cprBox.value = "" Then
    MsgBox "Du skal indsætte cprnummer", vbCritical
    Exit Sub
End If
 

' Mandatory field
If apikeyBox.value = "" Then
    MsgBox "Du skal indsætte api nøgle", vbCritical
    Exit Sub
End If

If SecureLifeButton.value = False And BedstPensionButton.value = False And SecureLifeTestButton.value = False Then
    MsgBox "Du skal vælge mellem bedstpension / securelife", vbCritical
    Exit Sub
End If

Set objRequest = CreateObject("MSXML2.XMLHTTP")

If SecureLifeButton = True Then
    Debug.Print ("Securelife")
    strUrl = "https://europe-west1-life-prod-e2f1e.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If

If BedstPensionButton = True Then
    Debug.Print ("Bedstpension")
    strUrl = "https://europe-west1-bedstpension-prod.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If


If SecureLifeTestButton = True Then
    Debug.Print ("Securelife test")
    strUrl = "https://europe-west1-life-stage-e2fb7.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If


Dim key As Variant
Dim value As Variant

Dim cellIndex As Integer
cellIndex = 6
' Adding data to Stamoplysninger sheet
For i = 6 To 23
    If cellIndex = 15 Then
        value = Cells(cellIndex, 3).value * 100
        key = Cells(cellIndex, 2).value
    ElseIf cellIndex = 16 Then
        value = Cells(cellIndex, 3).value * 100
        key = Cells(cellIndex, 2).value
    Else
        value = Cells(cellIndex, 3).value
        key = Cells(cellIndex, 2).value
    End If
    cellIndex = cellIndex + 1
    Debug.Print (key)
    Debug.Print (value)
    dict.Add key:=key, Item:=value
    If cellIndex = 13 Then
        cellIndex = cellIndex + 1 ' Skip row 13 (alder)
    End If
Next i

 Dim pensionSheet As Worksheet
 Dim pensionType As String
 Dim priceGroup As String
 ' Choose which sheet to fill data based on <pension type>
 pensionType = "AP Pension"

 Set pensionSheet = Sheets(pensionType)
 dict.Add key:="Frivilligt bidrag", Item:=pensionSheet.Cells(4, 3).value
dict.Add key:="Tab af erhvervsevne", Item:=pensionSheet.Cells(14, 2).value * 100
dict.Add key:="Invalidesum", Item:=pensionSheet.Cells(19, 2).value
dict.Add key:="Dødsfaldsdækning", Item:=pensionSheet.Cells(22, 2).value * 100
dict.Add key:="Børnerente", Item:=pensionSheet.Cells(26, 2).value
dict.Add key:="Kritisk sygdom", Item:=pensionSheet.Cells(29, 2).value
dict.Add key:="Kritisk sygdom til børn u. 21 år", Item:=pensionSheet.Cells(32, 2).value
dict.Add key:="Prisgruppe", Item:=pensionSheet.Cells(3, 2).value



Dim text As String
' Generate a text of key value pair for popup & printing the dict
For Each key In dict
    Dim val As Variant
    If IsNull(dict(key)) Then
        val = ""
    Else
        val = CStr(dict(key))
    End If
    text = text + " " + key + ": " + val + vbNewLine
Next

  
Dim answer As Integer
answer = MsgBox(text, vbQuestion + vbYesNo + vbDefaultButton2, "Import af medarbejderen")

End Sub




Private Sub ImportButton_Click()



Dim JsonObject As Object
Dim objRequest As Object
Dim strUrl As String
Dim blnAsync As Boolean
Dim strResponse As String
Dim codeResponse As String
Dim responseStatus As Variant


' Mandatory field
If cprBox.value = "" Then
    MsgBox "Du skal indsætte cprnummer", vbCritical
    Exit Sub
End If
 

' Mandatory field
If apikeyBox.value = "" Then
    MsgBox "Du skal indsætte api nøgle", vbCritical
    Exit Sub
End If

If SecureLifeButton.value = False And BedstPensionButton.value = False And SecureLifeTestButton.value = False Then
    MsgBox "Du skal vælge mellem bedstpension / securelife", vbCritical
    Exit Sub
End If

Set objRequest = CreateObject("MSXML2.XMLHTTP")

If SecureLifeButton = True Then
    Debug.Print ("Securelife")
    strUrl = "https://europe-west1-life-prod-e2f1e.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If

If BedstPensionButton = True Then
    Debug.Print ("Bedstpension")
    strUrl = "https://europe-west1-bedstpension-prod.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If


If SecureLifeTestButton = True Then
    Debug.Print ("Securelife test")
    strUrl = "https://europe-west1-life-stage-e2fb7.cloudfunctions.net/employeePolicy/export/" & cprBox.value
End If


blnAsync = True

With objRequest
    .Open "GET", strUrl, blnAsync
    .setRequestHeader "Content-Type", "application/json"
    .setRequestHeader "apikey", apikeyBox.value
    .send
    'spin wheels whilst waiting for response
    While objRequest.readyState <> 4
        DoEvents
    Wend
    strResponse = .responseText
    responseStatus = .Status
End With
 Debug.Print (responseStatus)
 
 If Not responseStatus = 200 Then
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


Dim answer As Integer
answer = MsgBox(text, vbQuestion + vbYesNo + vbDefaultButton2, "Import af medarbejderen")
If answer = vbYes Then
    Dim cellIndex As Integer
    cellIndex = 6
    ' Adding data to Stamoplysninger sheet
    For i = 6 To 23
        If cellIndex = 15 Then
            Cells(cellIndex, 3).value = JsonObject(Cells(cellIndex, 2).value) / 100
        ElseIf cellIndex = 16 Then
            Cells(cellIndex, 3).value = JsonObject(Cells(cellIndex, 2).value) / 100
        Else
            Cells(cellIndex, 3).value = JsonObject(Cells(cellIndex, 2).value)
        End If
        cellIndex = cellIndex + 1
        If cellIndex = 13 Then
            cellIndex = cellIndex + 1 ' Skip row 13 (alder)
        End If
    Next i
    
    Dim pensionSheet As Worksheet
    Dim pensionType As String
    Dim priceGroup As String
    ' Choose which sheet to fill data based on <pension type>
    pensionType = JsonObject("Pension type")
   
    sheetExist (pensionType)

    Set pensionSheet = Sheets(pensionType)
    pensionSheet.Cells(4, 3).value = JsonObject("Frivilligt bidrag") / 100
    pensionSheet.Cells(14, 2).value = JsonObject("Tab af erhvervsevne") / 100
    pensionSheet.Cells(19, 2).value = JsonObject("Invalidesum")
    pensionSheet.Cells(22, 2).value = JsonObject("Dødsfaldsdækning") / 100
    pensionSheet.Cells(26, 2).value = JsonObject("Børnerente")
    pensionSheet.Cells(29, 2).value = JsonObject("Kritisk sygdom")
    pensionSheet.Cells(32, 2).value = JsonObject("Kritisk sygdom til børn u. 21 år")
    pensionSheet.Cells(3, 11).value = JsonObject("Prisgruppe")
    MsgBox "Import OK"
    UserForm1.Hide
Else
  MsgBox "Import fejlede"
End If


End Sub

