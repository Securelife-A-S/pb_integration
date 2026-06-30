# Aktivér automatiske opdateringer (engangsopsætning)

Arket opdaterer sig selv, når du åbner det. For at det kan lykkes, skal Excel have lov til at
opdatere sine egne makroer. Får du fejlen **"Automatisk adgang til Visual Basic-projektet er ikke
pålidelig"** (eller bliver du ved med at have en gammel version), mangler denne indstilling.

## Sådan slår du den til (ét minut)

1. **Filer** → **Indstillinger**
2. **Sikkerhedscenter** → knappen **Indstillinger for Sikkerhedscenter…**
3. **Indstillinger for makroer**
4. Sæt flueben i **"Hav tillid til adgang til VBA-projektobjektmodellen"**
5. **OK**

## Hvis du allerede sidder fast på en gammel version

1. Slå indstillingen til som ovenfor.
2. **Luk Excel helt.**
3. Slet mappen **`pb`** på dit **Skrivebord**.
4. Åbn arket igen — opdateringen kører nu forfra og henter den nye version.

## Til IT (udrulning på mange maskiner)

Indstillingen svarer til registry-værdien `AccessVBOM`. Kan sættes via GPO eller:

```
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f
```

(`16.0` = Microsoft 365 / Office 2016 / 2019 / 2021.)
