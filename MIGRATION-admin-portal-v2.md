# Migrering til admin-portal-v2 (portal-token)

Den gamle apikey-flow mod slportal (Firebase cloud functions) er udfaset. Makroen henter nu
data fra **admin-portal-v2** med et **portal-token** (Sanctum Bearer-token).

## Hvad er ændret

Kun `UserForm1.frm` (`ImportButton_Click`):

- URL: `…cloudfunctions.net/integration/excel/{cpr}` → `{BASE_URL}/api/excel/employee/{cpr}`.
- Auth-header: `apikey: <key>` → `Authorization: Bearer <token>`.
- Det eksisterende felt (`apikeyBox`) genbruges — brugeren indsætter nu et **portal-token** i stedet
  for den gamle apikey. **Ingen ændringer i formularens kontroller** (ingen `.frx`-ændring nødvendig).
- `ExportButton_Click` er deaktiveret (v2 er skrivebeskyttet, intet PUT-endpoint).

`pensionBrokerExport.bas` er **uændret**.

## Sådan får rådgiveren et token

1. Log ind i portalen (portal.cpof.dk) som rådgiver.
2. Gå til **Indstillinger → API-nøgler** (`/settings/api-tokens`).
3. Opret et token (vælg fx 365 dages udløb), kopiér det (vises kun én gang).
4. Indsæt det i token-feltet i Excel-formularen.

Tokenet udløber (max 1 år) — når det sker, giver opslaget en fejl (401); generér da blot et nyt
token i portalen og indsæt det igen. Samme arbejdsgang som med den gamle nøgle.

## Miljø-mapping

| Radioknap | URL |
|---|---|
| SecureLife (prod) | `https://portal.cpof.dk` (CPOF — aktiv bruger) |
| SecureLife test | `https://test2.portal.cpof.dk` *(bekræft public staging-URL i koden)* |
| BedstPension | ikke migreret → viser besked |

## Anbefalet (valgfrit) — relabel feltet

Funktionelt virker det med det nuværende felt-navn/label. For klarhed kan label-teksten ændres fra
"Api-nøgle" til **"Portal-token"** i VBA-editoren (kontrollen ligger i den binære `.frx`, så det
kræver Excel — men det er kun kosmetisk og ikke nødvendigt for at det virker).

## Udrulning

1. Test mod prod med et rigtigt portal-token (kendt CPR udfylder "Portal import-eksport").
2. Commit `UserForm1.frm`.
3. **Bump `version.txt`** (0.3.0-alpha → næste) — trigger auto-update for alle rådgivere ved næste
   workbook-åbning (`init.bas`).
4. Merge til `main`.

## Test (jf. admin-portal/docs/integrations/pb-integration-update.md)

1. Indsæt portal-token + kendt CPR → "Portal import-eksport" udfyldes som før.
2. Forkert/udløbet token → 401-fejl vises (generér nyt token).
3. CPR uden adgang → "Du har ikke adgang til denne medarbejders virksomhed."
4. Bekræft at alle ark-felter udfyldes som før (uændret nøgle-mapping).
