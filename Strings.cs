namespace OutlookClassicSearch;

internal static class Strings
{
    public static bool IsEnglish { get; set; }

    // --- Menu ---
    public static string MenuSettings   => IsEnglish ? "⚙ Settings"    : "⚙ Instellingen";
    public static string MenuHelp       => "❓ Help";
    public static string MenuHelpItem   => "Help";
    public static string MenuInfoItem   => "Info";

    // --- Main form ---
    public static string LabelQuery      => IsEnglish ? "Search term:"  : "Zoekterm:";
    public static string QueryPlaceholder => IsEnglish ? "E.g. invoice, order number, customer name" : "Bijv. factuur, ordernummer, klantnaam";
    public static string BtnSearch       => IsEnglish ? "Search"        : "Zoeken";
    public static string BtnCancel       => IsEnglish ? "Cancel"        : "Annuleren";
    public static string BtnSave         => IsEnglish ? "Save"          : "Opslaan";
    public static string BtnClose        => IsEnglish ? "Close"         : "Sluiten";
    public static string ChkShowPreview  => IsEnglish ? "Show preview"  : "Toon voorbeeldvenster";
    public static string ChkUseDateRange => IsEnglish ? "Use date range": "Gebruik datum";

    // --- Filter placeholders ---
    public static string FilterMailbox    => "Filter: Mailbox";
    public static string FilterFolder     => IsEnglish ? "Filter: Folder"     : "Filter: Map";
    public static string FilterDate       => IsEnglish ? "Filter: Date"        : "Filter: Datum";
    public static string FilterSubject    => IsEnglish ? "Filter: Subject"     : "Filter: Onderwerp";
    public static string FilterSender     => IsEnglish ? "Filter: Sender"      : "Filter: Afzender";
    public static string FilterRecipients => IsEnglish ? "Filter: Recipients"  : "Filter: Geadresseerde";
    public static string FilterHasAttachment => IsEnglish ? "Has attachment"   : "Bijlage";
    public static string ColHasAttachment    => IsEnglish ? "Att."             : "Bijlage";
    public static string FilterYes           => IsEnglish ? "Yes" : "Ja";
    public static string FilterNo            => IsEnglish ? "No"  : "Nee";

    // --- Preview panel ---
    public static string BtnClosePreview => IsEnglish ? "Close preview" : "Venster sluiten";
    public static string BtnCopyPreview  => IsEnglish ? "Copy preview"  : "Kopieer voorbeeld";

    // --- Column headers ---
    public static string ColMailbox    => "Mailbox";
    public static string ColFolder     => IsEnglish ? "Folder"      : "Map";
    public static string ColDate       => IsEnglish ? "Date"         : "Datum";
    public static string ColSubject    => IsEnglish ? "Subject"      : "Onderwerp";
    public static string ColSender     => IsEnglish ? "Sender"       : "Afzender";
    public static string ColRecipients => IsEnglish ? "Recipients"   : "Geadresseerde";

    // --- Status bar ---
    public static string StatusReady           => IsEnglish ? "Ready for search"      : "Klaar voor zoekopdracht";
    public static string StatusSearchStarted   => IsEnglish ? "Search started..."     : "Zoeken gestart...";
    public static string StatusSearchDoneFmt   => IsEnglish ? "Done. {0} result(s)."  : "Klaar. {0} resultaat/resultaten.";
    public static string StatusSearchCancelled => IsEnglish ? "Search cancelled."     : "Zoeken geannuleerd.";
    public static string StatusSearchFailed    => IsEnglish ? "Search failed."         : "Zoeken mislukt.";
    public static string StatusCopied          => IsEnglish ? "Preview copied to clipboard." : "Voorbeeld gekopieerd naar klembord.";
    public static string StatusCopyFailed      => IsEnglish ? "Copy failed."            : "Kopieren mislukt.";
    public static string StatusAutoIndexDone   => IsEnglish ? "Auto-index updated."     : "Auto-index bijgewerkt.";
    public static string StatusAutoIndexPrefix => IsEnglish ? "Auto-index: "            : "Auto-index: ";
    public static string StatusAutoIndexFailed => IsEnglish ? "Auto-index failed."       : "Auto-index mislukt.";
    public static string StatusMailboxesFoundFmt  => IsEnglish ? "{0} mailbox(es) found." : "{0} mailbox(en) gevonden.";
    public static string StatusMailboxesFailed    => IsEnglish ? "Loading mailboxes failed." : "Mailboxen laden mislukt.";

    // --- Error dialogs (Form1) ---
    public static string MsgEnterQuery         => IsEnglish ? "Please enter a search term."           : "Vul een zoekterm in.";
    public static string MsgEnterQueryTitle    => IsEnglish ? "Search term missing"                   : "Zoekterm ontbreekt";
    public static string MsgSelectMailbox      => IsEnglish ? "Please select at least 1 mailbox."    : "Selecteer minimaal 1 mailbox.";
    public static string MsgSelectMailboxTitle => IsEnglish ? "Mailbox missing"                       : "Mailbox ontbreekt";
    public static string ErrSearchTitle        => IsEnglish ? "Search failed"                          : "Zoeken mislukt";
    public static string ErrSearchSummary      => IsEnglish ? "Search in Outlook failed."              : "Zoeken in Outlook is mislukt.";
    public static string ErrSearchHint         => IsEnglish ? "Check that Outlook Classic is running and that your mailboxes are accessible." : "Controleer of Outlook Classic draait en of je mailboxen bereikbaar zijn.";
    public static string ErrOpenMailTitle      => IsEnglish ? "Opening mail failed"                   : "Mail openen mislukt";
    public static string ErrOpenMailSummary    => IsEnglish ? "The selected mail could not be opened." : "De geselecteerde mail kon niet worden geopend.";
    public static string ErrOpenMailHint       => IsEnglish ? "The mail may have been deleted or you no longer have access to this mailbox." : "De mail is mogelijk verwijderd of je hebt geen rechten meer op deze mailbox.";
    public static string ErrMailboxTitle       => IsEnglish ? "Loading mailboxes failed"               : "Mailboxen laden mislukt";
    public static string ErrMailboxSummary     => IsEnglish ? "Mailboxes could not be retrieved from Outlook." : "Mailboxen konden niet uit Outlook worden opgehaald.";
    public static string ErrMailboxHint        => IsEnglish ? "Start Outlook Classic manually, check your profile and try again." : "Start Outlook Classic handmatig, controleer je profiel en probeer daarna opnieuw.";
    public static string ErrCopyTitle          => IsEnglish ? "Copy failed"                            : "Kopieren mislukt";
    public static string ErrCopySummary        => IsEnglish ? "The preview could not be copied."       : "Het voorbeeld kon niet worden gekopieerd.";

    // --- Settings form ---
    public static string SettingsTitle               => IsEnglish ? "Settings"                                  : "Instellingen";
    public static string SettingsChkSearchBody       => IsEnglish ? "Search also in body"                       : "Zoek ook in body";
    public static string SettingsChkSearchAtt        => IsEnglish ? "Search also in attachments"                : "Zoek ook in bijlagen";
    public static string SettingsChkUseIndex         => IsEnglish ? "Use persistent index"                      : "Gebruik permanente index";
    public static string SettingsLblMaxResults       => IsEnglish ? "Max results:"                              : "Max regels:";
    public static string SettingsChkExcludeExt       => IsEnglish ? "Exclude attachment extensions (;):"        : "Bijlage-extensies uitsluiten (;):";
    public static string SettingsLblStores           => IsEnglish ? "Mailboxes (incl. shared):"                 : "Mailboxen (incl. shared):";
    public static string SettingsBtnRefreshStores    => IsEnglish ? "Refresh mailboxes"                         : "Mailboxen vernieuwen";
    public static string SettingsBtnExcludeFolders   => IsEnglish ? "Exclude folders from search..."            : "Mappen uitsluiten van zoeken...";
    public static string SettingsLblExcludedFoldersFmt => IsEnglish ? "Excluded folders: {0}"                   : "Uitgesloten mappen: {0}";
    public static string SettingsBtnManageIndex      => IsEnglish ? "Manage index..."                           : "Index beheren...";
    public static string SettingsLblLanguage         => IsEnglish ? "Display language:"                         : "Weergavetaal:";
    public static string SettingsLblSearchHistoryMax => IsEnglish ? "Max search history:"                       : "Max zoekgeschiedenis:";
    public static string SettingsBtnClearHistory     => IsEnglish ? "Clear history"                             : "Geschiedenis wissen";
    public static string SettingsMsgNoMailbox        => IsEnglish ? "Please select at least 1 mailbox first."  : "Selecteer eerst minimaal 1 mailbox.";
    public static string SettingsMsgNoMailboxTitle   => IsEnglish ? "No mailboxes"                              : "Geen mailboxen";
    public static string SettingsMsgLoadingMailboxes => IsEnglish ? "Loading mailboxes..."                      : "Mailboxen ophalen...";
    public static string SettingsMsgMailboxesFailed  => IsEnglish ? "Loading failed."                           : "Laden mislukt.";
    public static string SettingsMsgLoadingFolders   => IsEnglish ? "Loading folder structure..."               : "Mappenstructuur laden...";
    public static string ErrMailboxSettingsTitle     => IsEnglish ? "Loading mailboxes failed"                  : "Mailboxen laden mislukt";
    public static string ErrMailboxSettingsSummary   => IsEnglish ? "Mailboxes could not be retrieved."         : "Mailboxen konden niet worden opgehaald.";

    // --- Index manager ---
    public static string IndexTitle              => IsEnglish ? "Index manager"                 : "Indexbeheer";
    public static string IndexScopeNotLoaded     => IsEnglish ? "Mailbox scope: not loaded"    : "Mailboxscope: niet geladen";
    public static string IndexScopeSelectedFmt   => IsEnglish ? "Mailbox scope: {0} selected"  : "Mailboxscope: {0} geselecteerd";
    public static string IndexStateNone          => IsEnglish ? "Index: not yet built"          : "Index: nog niet opgebouwd";
    public static string IndexStateFmt           => IsEnglish ? "Index: {0} items, last: {1}"  : "Index: {0} items, laatst: {1}";
    public static string IndexBuildProgressFmt   => IsEnglish ? "Index: {0}"                   : "Index: {0}";
    public static string IndexBtnSelectFolders   => IsEnglish ? "Choose folders for index"     : "Mappen voor index kiezen";
    public static string IndexBtnRefresh         => IsEnglish ? "Refresh index now"             : "Index nu verversen";
    public static string IndexBtnClear           => IsEnglish ? "Delete index"                  : "Index verwijderen";
    public static string IndexChkAutoRefresh     => IsEnglish ? "Auto refresh"                  : "Automatisch verversen";
    public static string IndexLblInterval        => IsEnglish ? "Interval (min):"               : "Interval (min):";
    public static string IndexFolderTitle        => IsEnglish ? "Choose folders for index"      : "Mappen opnemen in index";
    public static string IndexErrTitle           => IsEnglish ? "Indexing failed"               : "Indexeren mislukt";
    public static string IndexErrSummary         => IsEnglish ? "Building the index failed."    : "Het opbouwen van de index is mislukt.";
    public static string IndexMsgNoMailbox       => IsEnglish ? "Please select mailboxes in the main screen first." : "Selecteer eerst mailboxen in het hoofdscherm.";
    public static string IndexMsgNoMailboxTitle  => IsEnglish ? "No mailboxes"                  : "Geen mailboxen";
    public static string IndexMsgBusy            => IsEnglish ? "Wait for indexing to finish before closing this window." : "Wacht tot de indexering klaar is voordat je dit venster sluit.";
    public static string IndexMsgBusyTitle       => IsEnglish ? "Indexing in progress"          : "Indexering actief";
    public static string IndexChkIndexOnStartup  => IsEnglish ? "Build index on startup"        : "Indexeren bij opstarten";

    // --- Folder selection ---
    public static string FolderCheckAndSave   => IsEnglish ? "Check folders and click Save."  : "Vink mappen aan en klik op Opslaan.";
    public static string FolderLoading        => IsEnglish ? "Loading folder structure..."    : "Mappenstructuur laden...";
    public static string FolderLoadCancelled  => IsEnglish ? "Loading cancelled."             : "Laden geannuleerd.";
    public static string FolderLoadFailed     => IsEnglish ? "Loading failed. Close this window and try again." : "Laden mislukt. Sluit dit venster en probeer opnieuw.";
    public static string FolderExcludeTitle   => IsEnglish ? "Exclude folders from search"   : "Mappen uitsluiten van zoeken";

    // --- Info / Help forms ---
    public static string InfoAppName    => "Outlook Classic Search";
    public static string InfoVersionFmt => IsEnglish ? "Version {0}"  : "Versie {0}";
    public static string InfoDesc       => IsEnglish
        ? "Search email messages in Outlook Classic quickly and clearly\nvia multiple mailboxes, with filters, preview pane and index support."
        : "Doorzoek e-mailberichten in Outlook Classic snel en overzichtelijk\nvia meerdere mailboxen, met filters, voorbeeldvenster en indexondersteuning.";
    public static string InfoBuiltWith  => IsEnglish
        ? "Built with .NET 8 / WinForms and NetOffice.OutlookApi.\n© 2026 – mjeppo"
        : "Ontwikkeld met .NET 8 / WinForms en NetOffice.OutlookApi.\n© 2026 – mjeppo";

    public static string HelpContent => IsEnglish
        ? "Searching\n" +
          "  • Enter a search term and press Enter or click Search.\n" +
          "  • Enable 'Use date range' to filter by a date period.\n" +
          "  • Use ⚙ Settings to choose mailboxes, folders and search options.\n\n" +
          "Results\n" +
          "  • Click a column header to sort; click again to reverse.\n" +
          "  • Drag a column header to reorder columns.\n" +
          "  • Right-click a column header to hide or show columns.\n" +
          "  • Type in a filter field above a column to filter search results.\n" +
          "  • Double-click a row to open the email in Outlook.\n\n" +
          "Preview pane\n" +
          "  • Check 'Show preview' to see a message preview.\n" +
          "  • Drag the splitter to resize the preview pane.\n" +
          "  • Use 'Copy preview' to copy the message text to the clipboard."
        : "Zoeken\n" +
          "  • Typ een zoekterm en druk op Enter of klik Zoeken.\n" +
          "  • Schakel 'Gebruik datum' in om te filteren op een datumperiode.\n" +
          "  • Via ⚙ Instellingen kies je mailboxen, mappen en zoekopties.\n\n" +
          "Resultaten\n" +
          "  • Klik op een kolomkop om te sorteren; nogmaals klikken keert de volgorde om.\n" +
          "  • Sleep een kolomkop om de volgorde van kolommen te wijzigen.\n" +
          "  • Rechtsklik op een kolomkop om kolommen te verbergen of weer te tonen.\n" +
          "  • Typ in een filterveld boven een kolom om de zoekresultaten direct te filteren.\n" +
          "  • Dubbelklik op een rij om het e-mailbericht in Outlook te openen.\n\n" +
          "Voorbeeldvenster\n" +
          "  • Vink 'Toon voorbeeldvenster' aan om een berichtvoorbeeld te zien.\n" +
          "  • Versleep de splitter om het voorbeeldvenster groter of kleiner te maken.\n" +
          "  • Gebruik 'Kopieer voorbeeld' om de berichttekst naar het klembord te kopiëren.";
}
