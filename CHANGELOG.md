# Changelog

Alle belangrijke wijzigingen in dit project worden gedocumenteerd in dit bestand.

## [1.5.0] - 2025-01-XX

### Toegevoegd
- **Toepassen knop** in Instellingenscherm: Wijzigingen kunnen nu worden toegepast zonder het scherm te sluiten
- **Mappenfilter dropdown**: Zoeken kan nu beperkt worden tot specifieke mappen via een selecteerbare mappenstructuur
- **"Clear filters" knop**: Wordt getoond wanneer er een map is geselecteerd voor filtering
- **Achtergrond indexeren**: Het indexeren kan nu op de achtergrond doorlopen terwijl je verder werkt
  - Geanimeerde statusbalk indicator toont wanneer indexeren actief is
  - Bevestigingsdialoog bij sluiten van index manager tijdens actief indexeren
- **Waarschuwingen voor zoekresultaten**:
  - "Geen resultaten" melding wanneer zoekopdracht niets oplevert
  - "Te veel resultaten" waarschuwing met instelbare drempelwaarde (standaard 110)
  - Drempelwaarde configureerbaar in Instellingen (10-10.000 bereik)
- **CSV Export**: Zoekresultaten kunnen nu geëxporteerd worden naar CSV-bestand
- **Versienummer in titelbalk**: Het versienummer wordt nu weergegeven in de applicatie titelbalk
- **Mappenstructuur preloading**: Mappenstructuur wordt bij opstarten geladen voor snellere filtering

### Verbeterd
- **Menu-indeling**: Export en Index functies hebben nu aparte menu's voor betere organisatie
- **Dark mode MenuStrip hover**: Hover-achtergrondkleur in dark mode aangepast naar donkergrijs voor betere leesbaarheid (was lichtblauw met witte tekst)
- **Folder filter dropdown UX**: Dropdown is nu een proper dialog met close knop en ESC/Enter shortcuts
- **Statusberichten**: Duidelijke feedback bij laden van mappenstructuur bij opstarten

### Opgelost
- Mappenstructuur wordt niet meer bij elke klik opnieuw geladen (performance fix)
- Folder filter dropdown kan nu correct worden gesloten
- Taal en thema wijzigingen worden nu direct toegepast bij gebruik van "Toepassen" knop

### Technisch
- MaterialSwitch dark mode initialisatie bekend probleem gedocumenteerd (visueel, geen functionaliteit impact)
- SearchableFilterDropdown gewijzigd van borderless popup naar FormBorderStyle.FixedSingle
- IndexManagerForm gewijzigd van modal ShowDialog() naar modeless Show()
- BackgroundIndexingChanged event toegevoegd voor statusbar updates
- Timer-based animation (500ms interval) voor achtergrond indexeren indicator
- HashSet<string> _selectedSearchFolders voor folder filtering performance

## [1.4.0] - 2024-XX-XX

### Toegevoegd
- CSV export functionaliteit
- Versienummer in titelbalk

### Verbeterd
- Menu structuur geoptimaliseerd
- UI/UX verbeteringen

## [1.3.0] - 2024-XX-XX

### Toegevoegd
- Dark mode ondersteuning
- Theme selector (Light/Dark/Auto)
- Meertalige ondersteuning (Nederlands/Engels)

## [1.2.0] - 2024-XX-XX

### Toegevoegd
- Permanente index functionaliteit
- Index beheer venster
- Automatisch index verversen

## [1.1.0] - 2024-XX-XX

### Toegevoegd
- Zoeken in bijlagen
- Map exclusion functionaliteit

## [1.0.0] - 2024-XX-XX

### Toegevoegd
- Initiële release
- Basis zoekfunctionaliteit in Outlook Classic
- Ondersteuning voor gedeelde mailboxen
- Live zoekresultaten
- Datum filtering
