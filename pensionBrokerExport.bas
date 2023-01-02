Sub ExportToPensionBroker()

' ### Step 1: Load data from pension broker sheet and parse to dict ###
Dim key As Variant
Dim value As Variant
Dim dict As New Scripting.Dictionary
Dim startRow As Integer
Dim endRow As Integer
Dim xmlPath As String
Dim chromePath As String
Dim requestType As String

Dim strPath As String
strPath = Environ("USERPROFILE") & "\Desktop\pb"
startRow = 6 'Start row in Pension broker sheet
endRow = 30 'End row in Pension broker sheet


' xmlPath = "C:\Users\Andreas Styltsvig\Desktop\pensionbroker_integration\xml\"
xmlPath = strPath & "\pb_integration-main\xml\"

For i = startRow To endRow
    key = ThisWorkbook.Worksheets("Pension Broker").Range("B" & i).value
    value = ThisWorkbook.Worksheets("Pension Broker").Range("C" & i).value
    If InStr(value, ",") > 0 Then
    Debug.Print key
    Debug.Print value
    value = value * 100
    Debug.Print value
    value = Replace(value, ",", ".")
    Debug.Print value
    End If
    dict.Add key:=key, Item:=value
Next i

dict.Add key:="Løntype", Item:=ThisWorkbook.Worksheets("Pension Broker").Range("E25").value

For Each key In dict.keys
    Debug.Print key, dict(key)
Next key


' ### Step 2: Figure out which pensionCase ###
Dim pensionCompanyName As String, pensionType As String, pensionCase As String

' <-- We have to figure out what the pensioncase is named, to switch to correct name for pension broker
pensionCompanyName = dict("Pensionsselskab") '<--- Pensionsselskab navn

Select Case pensionCompanyName
    Case "AP Pension"
        ' APPensionPensionCase
        pensionCase = "APPensionPensionCase"
    Case "Euro Accident Liv"
        ' EuroAccidentCompanyPensionCase
        pensionCase = "EuroAccidentCompanyPensionCase"
    Case "Danica Pension"
        ' DanicaPensionCase
        pensionCase = "DanicaPensionCase"
    Case "Velliv"
        ' VellivETSPensionCase,VellivLandmandspensionPensionCase, VellivLivPensionCase, VellivN16PensionCase
        pensionType = dict("Produkt Velliv") '<--- Pensionsselskab navn
        If StrComp(pensionType, "Velliv N16") = 0 Then
            pensionCase = "VellivN16PensionCase"
        ElseIf StrComp(pensionType, "Velliv Landmandspension") = 0 Then
            pensionCase = "VellivLandmandspensionPensionCase"
        ElseIf StrComp(pensionType, "Velliv") = 0 Then
            pensionCase = "VellivLivPensionCase"
        ElseIf StrComp(pensionType, "Velliv ETS") = 0 Then
            pensionCase = "VellivETSPensionCase"
        End If
    Case "Topdanmark A/S"
        ' TopdanmarkCompanyExecutivePensionCase, TopdanmarkCompanyIndividualPensionCase,
        ' TopdanmarkCompanyPensionPensionCase. TopdanmarkCompanyProprietorPensionCase,
        ' TopdanmarkCompanyPseudoPrivatePensionCase
        pensionType = dict("Produkt Topdanmark A/S") '<--- Pensionsselskab navn
        If StrComp(pensionType, "FirmaPension") = 0 Then
            pensionCase = "TopdanmarkCompanyPensionPensionCase"
        ElseIf StrComp(pensionType, "Individuel firmaordning Profilpension/Link/Spar Top") = 0 Then
            pensionCase = "TopdanmarkCompanyPseudoPrivatePensionCase"
        ElseIf StrComp(pensionType, "Dirketørpension") = 0 Then
            pensionCase = "TopdanmarkCompanyExecutivePensionCase"
        ElseIf StrComp(pensionType, "Indehaverpension/Privatpension") = 0 Then
            pensionCase = "TopdanmarkCompanyProprietorPensionCase"
        ElseIf StrComp(pensionType, "Privatpension") = 0 Then
            pensionCase = "TopdanmarkCompanyIndividualPensionCase"
        End If
    Case "PFA Pension"
        pensionType = dict("Produkt PFA Pension")
        ' PFAPlusPensionCase , PFAKontantpensionPensionCase
        If StrComp(pensionType, "PFA Plus") = 0 Then
           pensionCase = "PFAPlusPensionCase"
        ElseIf StrComp(pensionType, "PFA Kontantpension") = 0 Then
          pensionCase = "PFAKontantpensionPensionCase"
        End If
End Select

Dim xmlDoc As Object, xmlRoot As Object, pathToXML As String

Set xmlDoc = CreateObject("MSXML2.DOMDocument")

pathToXML = xmlPath & pensionCase & ".xml" '<--- Path to the file

Debug.Print pathToXML

Dim tmpPath As String

tmpPath = xmlPath & "tmp.xml"

FileCopy pathToXML, tmpPath

Debug.Print tmpPath

Call xmlDoc.Load(tmpPath) ' <-- Load file


' Pensioncase
Set xmlRoot = xmlDoc.getElementsByTagName("InputData").Item(0)

xmlRoot.SelectSingleNode("PensionCase").setAttribute("xsi:type") = pensionCase
Call xmlDoc.Save(tmpPath)

' CVR & CPR
Set xmlRoot = xmlDoc.getElementsByTagName("PensionCase").Item(0)
myVar = dict("CVR nr.") '<--- Fornavn
xmlRoot.SelectSingleNode("CVR").text = myVar
myVar = dict("CPR nr.") '<--- Efternavn
xmlRoot.SelectSingleNode("CPR").text = myVar
myVar = dict("Type af begæring")

If InStr(myVar, "Nytegning") > 0 Then
myVar = "Subscription"
End If

If InStr(myVar, "Ændring") > 0 Then
myVar = "Amendment"
End If

xmlRoot.SelectSingleNode("RequestType").text = myVar
Call xmlDoc.Save(tmpPath)



' PensionTaker / Personlig info
Set xmlRoot = xmlDoc.getElementsByTagName("PensionTaker").Item(0)

myVar = dict("Fornavn") '<--- Fornavn
xmlRoot.SelectSingleNode("FirstName").text = myVar
myVar = dict("Efternavn") '<--- Efternavn
xmlRoot.SelectSingleNode("LastName").text = myVar
myVar = dict("Telefon") '<--- Telefon
xmlRoot.SelectSingleNode("TelephoneNo1").text = myVar
myVar = dict("E-mail") '<--- Email
xmlRoot.SelectSingleNode("Email").text = myVar
myVar = dict("Virksomhedsnavn") '<--- Virksomhedsnavn
xmlRoot.SelectSingleNode("EmployerCompanyName").text = myVar

Call xmlDoc.Save(tmpPath)


Call xmlDoc.Load(tmpPath) ' <-- Load file

' Contribution / Bidrag
Set xmlRoot = xmlDoc.getElementsByTagName("Contribution").Item(0)

myVar = dict("Løn") '<--- Løn
xmlRoot.SelectSingleNode("AnnualSalary").text = myVar

Debug.Print (pensionCase)
' today = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "YYYY-DD-MM")
' Debug.Print today

Select Case pensionCase
    Case "APPensionPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployerContribution").text = myVar
        myVar = dict("Obligatorisk medarbejderbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployeeContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        
    Case "PFAPlusPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployerContribution").text = myVar
        myVar = dict("Obligatorisk medarbejderbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployeeContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        
    Case "DanicaPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
        xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        
    Case "EuroAccidentCompanyPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployerContribution").text = myVar
        
        
    Case "VellivN16PensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
        xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        myVar = dict("Løn")
        xmlRoot.SelectSingleNode("BonusSalary").text = myVar
        If dict("Frivilligtbidrag") > 0 Then
            xmlRoot.SelectSingleNode("PremiumWaiver").text = "True"
        Else
             Debug.Print "Deleting premiumwaiver"
            Set deleteMe = xmlRoot.SelectSingleNode("PremiumWaiver")
            Set oldChild = deleteMe.ParentNode.RemoveChild(deleteMe)
        End If
         
    Case "VellivLandmandspensionPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployerContribution").text = myVar
        myVar = dict("Obligatorisk medarbejderbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployeeContribution").text = myVar
        ' No optional contribution
        
    Case "VellivETSPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployerContribution").text = myVar
        myVar = dict("Obligatorisk medarbejderbidrag")
        xmlRoot.SelectSingleNode("MandatoryEmployeeContribution").text = myVar
        ' No optional contribution
        
    Case "VellivLivPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
        xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        
    Case "TopdanmarkCompanyPensionPensionCase"
        myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
        xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
        myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
        xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
        ' xmlRoot.SelectSingleNode("OptionalContributionStartDate").Text = today
        
         
    Case "TopdanmarkCompanyPseudoPrivatePensionCase"
         myVar = dict("Obligatorisk arbejdsgiverbidrag")
         xmlRoot.SelectSingleNode("EmployerContribution").text = myVar
         myVar = dict("Obligatorisk medarbejderbidrag")
         xmlRoot.SelectSingleNode("EmployeeContribution").text = myVar
         If dict("Frivilligtbidrag") > 0 Then
            xmlRoot.SelectSingleNode("PremiumWaiver").text = "True"
        Else
            Debug.Print "Deleting premiumwaiver"
            Set deleteMe = xmlRoot.SelectSingleNode("PremiumWaiver")
            Set oldChild = deleteMe.ParentNode.RemoveChild(deleteMe)
        End If
        
    Case "TopdanmarkCompanyIndividualPensionCase"
         myVar = dict("Obligatorisk arbejdsgiverbidrag")
         xmlRoot.SelectSingleNode("EmployerContribution").text = myVar
         myVar = dict("Obligatorisk medarbejderbidrag")
         xmlRoot.SelectSingleNode("EmployeeContribution").text = myVar
         If dict("Frivilligtbidrag") > 0 Then
            xmlRoot.SelectSingleNode("PremiumWaiver").text = "True"
        Else
            Debug.Print "Deleting premiumwaiver"
            Set deleteMe = xmlRoot.SelectSingleNode("PremiumWaiver")
            Set oldChild = deleteMe.ParentNode.RemoveChild(deleteMe)
        End If
    
    Case "TopdanmarkCompanyExecutivePensionCase" < --BROKEN
         myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
         xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
         myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
         xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
         If dict("Frivilligtbidrag") > 0 Then
            xmlRoot.SelectSingleNode("PremiumWaiver").text = "True"
        Else
            Debug.Print "Deleting premiumwaiver"
            Set deleteMe = xmlRoot.SelectSingleNode("PremiumWaiver")
            Set oldChild = deleteMe.ParentNode.RemoveChild(deleteMe)
        End If
         
    Case "TopdanmarkCompanyProprietorPensionCase" < --BROKEN
         myVar = dict("Obligatorisk arbejdsgiverbidrag") + dict("Obligatorisk medarbejderbidrag")  '<--- Samlet obligatorisk arbejdsgiver og medarbejderbidrag
         xmlRoot.SelectSingleNode("MandatoryContribution").text = myVar
         myVar = dict("Frivilligtbidrag") '<--- Frivilligtbidrag
         xmlRoot.SelectSingleNode("OptionalContribution").text = myVar
         If dict("Frivilligtbidrag") > 0 Then
            xmlRoot.SelectSingleNode("PremiumWaiver").text = "True"
        Else
            Debug.Print "Deleting premiumwaiver"
            Set deleteMe = xmlRoot.SelectSingleNode("PremiumWaiver")
            Set oldChild = deleteMe.ParentNode.RemoveChild(deleteMe)
        End If
         
    
   
End Select


Call xmlDoc.Save(tmpPath)

' Investment / Investering
' <---- Nothing to insert here so far --->

If Not InStr(pensionCase, "EuroAccidentCompanyPensionCase") > 0 Then
' Savings / Opsparing
Set xmlRoot = xmlDoc.getElementsByTagName("Savings").Item(0)
xmlRoot.SelectSingleNode("FirstSavingsType").text = "PensionAnnuity"
xmlRoot.SelectSingleNode("FirstSavingsToTaxAllowance").text = "True"
xmlRoot.SelectSingleNode("TheRestSavingsType").text = "LifeAnnuity"
Call xmlDoc.Save(tmpPath)


End If

' Coverage / Dækning
myVar = dict("Type af begæring")

If InStr(myVar, "Nytegning") > 0 Then
Set xmlRoot = xmlDoc.getElementsByTagName("Coverage").Item(0)

Select Case pensionCase
    Case "APPensionPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"

    Case "DanicaPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        myVar = dict("Løn")
        xmlRoot.SelectSingleNode("RewardingSalary").text = myVar
        myVar = dict("Løntype")
        If InStr(myVar, "Standard") > 0 Then
        xmlRoot.SelectSingleNode("WorkAbilityLossScales").text = "false"
        ElseIf InStr(myVar, "Lønskala") > 0 Then
        xmlRoot.SelectSingleNode("WorkAbilityLossScales").text = "True"
        End If
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
        
    Case "EuroAccidentCompanyPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathFixedAmount").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        
    Case "VellivN16PensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        myVar = dict("Løntype") '<--- Børnerente
        If InStr(myVar, "Standard") > 0 Then
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxType").text = "Standard"
        ElseIf InStr(myVar, "Lønskala") > 0 Then
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxType").text = "Scale"
        End If
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
        
        
    Case "VellivLandmandspensionPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        ' myVar = dict("Invalidesum") '<--- Invalidesum
        ' xmlRoot.SelectSingleNode("DisabilityFixedAmount").Text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
    Case "VellivETSPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        ' myVar = dict("Invalidesum") '<--- Invalidesum
        ' xmlRoot.SelectSingleNode("DisabilityFixedAmount").Text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
    Case "VellivLivPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
    Case "TopdanmarkCompanyPensionPensionCase"
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
    
    Case "TopdanmarkCompanyPseudoPrivatePensionCase" ' <-- BROKEN
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar

    Case "TopdanmarkCompanyExecutivePensionCase" ' <-- BROKEN
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
    Case "TopdanmarkCompanyProprietorPensionCase" ' <-- BROKEN
        myVar = dict("Dødsfald") '<--- Dødsfald
        xmlRoot.SelectSingleNode("DeathPercent").text = myVar
        myVar = dict("Tab af erhvervsevne") '<--- Tab af erhvervsevne
        xmlRoot.SelectSingleNode("WorkAbilityLossPercent").text = myVar
        myVar = dict("Invalidesum") '<--- Invalidesum
        xmlRoot.SelectSingleNode("DisabilityFixedAmount").text = myVar
        myVar = dict("Kritisk sygdom") '<--- Invalidesum
        xmlRoot.SelectSingleNode("CriticalDiseaseFixedAmount").text = myVar
        myVar = dict("Børnerente") '<--- Børnerente
        xmlRoot.SelectSingleNode("ChildPensionPercent").text = myVar
        xmlRoot.SelectSingleNode("WorkAbilityLossTaxCode").text = "TaxCode1"
        xmlRoot.SelectSingleNode("DeathTaxCode").text = "TaxCode5"
        
    Case "TopdanmarkCompanyIndividualPensionCase"
        ' Do nothing
        
    Case "PFAPlusPensionCase"
       ' Do nothing
        
End Select
Debug.Print "NYTEGNING"
Call xmlDoc.Save(tmpPath)
End If

myVar = dict("Type af begæring")

If InStr(myVar, "Ændring") > 0 Then
Debug.Print "ÆNDRINGSBEGÆRING"
Dim list As IXMLDOMNodeList
Set list = xmlDoc.SelectNodes("//InputData/PensionCase/Coverage")


For Each Item In list
  Debug.Print Item.HasChildNodes
  For Each childNode In Item.ChildNodes
    Debug.Print "Remove node from tmp coverage"
    Debug.Print childNode.BaseName & " " & childNode.text
    childNode.ParentNode.RemoveChild childNode
  Next
Next
Call xmlDoc.Save(tmpPath)
End If

' ### Step 3: Export data to pension broker ###
Dim fullURL As String

fullURL = """http://pensionbroker.dk/Client/Sirius.PensionBroker.Client.Shell.application?inputfile=" & tmpPath & """"
' fullURL = """http://pensionbrokerdemo.dk/Client/Sirius.PensionBroker.Client.Shell.application?inputfile=" & tmpPath & """"

Debug.Print fullURL

Shell IIf(Left(Application.OperatingSystem, 3) = "Win", "explorer ", "open ") & _
    fullURL

Debug.Print "DONE"
End Sub



