# Outlook Classic Search

Windows desktop app (WinForms) om te zoeken in Outlook Classic, inclusief gedeelde mailboxen die in het Outlook-profiel geladen zijn (Exchange Online / Microsoft 365).

## Functies

- Zoeken op onderwerp, afzender en ontvangers
- Optioneel zoeken in mail body
- Live resultaten tijdens het zoeken (rijen verschijnen direct)
- Permanente index op schijf met apart indexbeheer-venster
- Index opnemen/excluderen per map, plus handmatig verversen
- Automatisch index verversen op interval
- Datumfilter (van/tot)
- Selectie van mailboxen (ook shared mailboxen)
- Uitsluiten van willekeurige submappen via mailbox-mapboom
- Opslaan van geselecteerde mailboxen en mapkeuzes
- Zoeken in bijlagen en extensies uitsluiten (bijv. .zip;.png)
- Dubbelklik op resultaat om de mail in Outlook te openen
- Max aantal resultaten instelbaar

## Vereisten

- Windows 10/11
- Outlook Classic (desktop) geinstalleerd en geconfigureerd
- Mailboxen en shared mailboxen zichtbaar in Outlook-profiel
- .NET SDK 8

De app gebruikt NetOffice interop wrappers. Daardoor is er geen directe runtime-afhankelijkheid meer op Office PIA assembly `office, Version=15.0.0.0`.

## Build

```powershell
dotnet restore
dotnet build -c Release
```

## Publiceren (installeerbare map)

```powershell
./publish.ps1 -Architecture both
```

Output:

- `bin/Release/net8.0-windows/win-x86/publish/OutlookClassicSearch.exe`
- `bin/Release/net8.0-windows/win-x64/publish/OutlookClassicSearch.exe`

Gebruik `win-x86` als Outlook 32-bit is. Gebruik `win-x64` als Outlook 64-bit is.

Kopieer de volledige publish-map naar de doelmachine en start `OutlookClassicSearch.exe`.

## Probleemoplossing: "Mailboxen laden mislukt"

Als de foutdetails een `FileNotFoundException (0x80070002)` tonen, is dit meestal een Outlook COM laadprobleem.

Controleer:

- Outlook Classic staat echt geinstalleerd en opent met hetzelfde Windows-profiel
- Je gebruikt de juiste app-architectuur (`win-x86` voor 32-bit Outlook, `win-x64` voor 64-bit Outlook)
- Outlook-profiel is volledig geladen (start Outlook eerst handmatig)

## Opmerking over shared mailboxen

De app leest mailboxen via Outlook COM Stores. Dit betekent:

- Shared mailboxen moeten toegevoegd/automapped zijn in Outlook Classic
- Rechten moeten correct zijn in Exchange/Microsoft 365
- De app gebruikt hetzelfde account/profiel als Outlook op die pc
