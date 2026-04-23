# Outlook Classic Search v1.5.0 - Release Notes

**Release datum:** April 2026

## 🎉 Nieuwe functies

### Mappenfiltering
- **Zoek in specifieke mappen**: Je kunt nu zoeken binnen geselecteerde mappen via een nieuw mappenfilter dropdown menu
- **"Clear filters" knop**: Wordt automatisch getoond wanneer je mappen hebt geselecteerd voor filtering
- **Snelle mappenstructuur**: De mappenstructuur wordt bij opstarten geladen voor directe beschikbaarheid

### Achtergrond indexeren
- **Non-blocking indexeren**: Het indexeren kan nu op de achtergrond doorlopen terwijl je verder werkt
- **Statusbalk indicator**: Geanimeerde "Achtergrond indexeren..." indicator in de statusbalk
- **Bevestiging bij sluiten**: Waarschuwing wanneer je het index manager venster sluit tijdens actief indexeren

### Zoekresultaat verbeteringen
- **"Geen resultaten" melding**: Oranje waarschuwingslabel wanneer zoekopdracht niets oplevert
- **"Te veel resultaten" waarschuwing**: Melding wanneer er te veel resultaten zijn met suggestie om verder te filteren
- **Instelbare drempelwaarde**: De waarschuwingsdrempel is configureerbaar (10-10.000, standaard 110)

### Export functionaliteit
- **CSV Export**: Exporteer zoekresultaten naar CSV-bestand voor verdere analyse in Excel
- **Proper escaping**: Correct omgaan met quotes, newlines en speciale tekens

### UI/UX verbeteringen
- **"Toepassen" knop in Instellingen**: Wijzigingen kunnen nu toegepast worden zonder het scherm te sluiten
- **Direct zichtbare wijzigingen**: Thema en taal worden direct toegepast bij gebruik van "Toepassen" knop
- **Versienummer in titelbalk**: Het versienummer is nu zichtbaar in de applicatie titelbalk
- **Verbeterde menu-indeling**: Export en Index hebben nu aparte menu's voor betere organisatie

## 🐛 Opgeloste problemen

### Performance fixes
- Mappenstructuur wordt niet meer bij elke klik opnieuw geladen (was 10 seconden delay)
- Folder filter dropdown kan nu correct worden gesloten (was freeze issue)

### Dark mode fixes
- MenuStrip hover achtergrondkleur in dark mode aangepast naar donkergrijs voor betere leesbaarheid
  - Was: lichtblauw met witte tekst (onleesbaar)
  - Nu: donkergrijs met witte tekst (goed leesbaar)

### Settings dialog
- Taal wijzigingen worden nu direct toegepast bij "Toepassen" knop
- Thema wijzigingen worden nu correct op het gehele formulier toegepast

## 📋 Technische details

### Architectuur wijzigingen
- `SearchableFilterDropdown`: Gewijzigd van borderless popup naar `FormBorderStyle.FixedSingle` met title en close button
- `IndexManagerForm`: Gewijzigd van modal `ShowDialog()` naar modeless `Show()` voor achtergrond operatie
- `BackgroundIndexingChanged` event toegevoegd voor real-time status updates
- Timer-based animation (500ms interval) voor statusbalk indicator

### Data structuren
- `HashSet<string> _selectedSearchFolders`: Voor snelle folder filtering lookup
- `SearchCriteria.IncludedFolderPaths`: Nieuwe property voor folder scope filtering

### Bekende beperkingen
- MaterialSwitch controls kunnen in dark mode bij eerste start korte donkere outlines tonen
  - Dit is een visueel issue zonder functionaliteit impact
  - Lost zichzelf op bij eerste interactie met dialogs
  - Geen performance impact

## 📦 Installatie

### Vereisten
- Windows 10/11
- Outlook Classic (desktop) geïnstalleerd en geconfigureerd
- .NET 8 Desktop Runtime (wordt automatisch aangeboden tijdens installatie indien niet aanwezig)

### Installatie bestanden
Kies de juiste installer voor jouw Outlook versie:

**64-bit Outlook (meest voorkomend):**
- `OutlookClassicSearch-v1.5.0-Setup-x64.exe` (~4.9 MB)

**32-bit Outlook:**
- `OutlookClassicSearch-v1.5.0-Setup-x86.exe` (~4.9 MB)

### Upgrade van eerdere versies
- Oude versie wordt automatisch verwijderd
- Instellingen en index blijven behouden
- Geen handmatige stappen nodig

## 🔄 Migratie

Als je upgrade vanaf versie 1.4.0 of eerder:
- Alle bestaande instellingen blijven behouden
- Bestaande index blijft werken
- Uitgesloten mappen blijven uitgesloten
- Mailbox selecties blijven bewaard
- Zoekgeschiedenis blijft beschikbaar

## 📝 Volgende versie (roadmap)

Mogelijke features voor v1.6.0:
- Geavanceerde zoekfilters (CC, BCC)
- Opgeslagen zoekopdrachten
- Export naar Excel (XLSX)
- Performance optimalisaties voor grote mailboxen

## 🐛 Bug reports & feedback

Meld bugs of feature requests via:
- GitHub Issues: https://github.com/mjeppo/outlook-classic-zoek/issues

---

**Veel zoekplezier met versie 1.5.0! 🚀**
