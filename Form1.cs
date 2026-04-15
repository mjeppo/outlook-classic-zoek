using System.ComponentModel;

namespace OutlookClassicSearch;

public partial class Form1 : Form
{
    private readonly BindingSource _resultsBindingSource = new();
    private readonly BindingList<EmailSearchResult> _visibleResults = new();
    private readonly List<EmailSearchResult> _allResults = new();
    private readonly AppSettings _settings;
    private readonly System.Windows.Forms.Timer _indexRefreshTimer = new();
    private readonly System.Windows.Forms.Timer _filterDebounceTimer = new() { Interval = 300 };

    private CancellationTokenSource? _searchCancellation;
    private CancellationTokenSource? _previewCancellation;
    private bool _isAutoIndexRunning;
    private List<MailboxFolderRoot> _folderRoots = new();
    private List<string> _folderTreeStoreIds = new();
    private List<StoreListItem> _availableStores = new();
    private Dictionary<string, TextBox> _columnFilterMap = new();
    private Dictionary<string, Button> _dropdownButtonMap = new();
    private Dictionary<string, HashSet<string>> _dropdownFilterMap = new();
    private string _sortColumn = nameof(EmailSearchResult.ReceivedTime);
    private bool _sortAscending;

    public Form1()
    {
        InitializeComponent();
        AppVisualAssets.ApplyWindowIcon(this);
        AppTheme.Apply(this);
        AppTheme.ApplyPrimaryStyle(btnSearch);

        _settings = AppSettingsStore.Load();
        Strings.IsEnglish = _settings.Language == "en";
        ApplyStrings();
        InitializeResultsGridColumns();
        _resultsBindingSource.DataSource = _visibleResults;
        dgvResults.DataSource = _resultsBindingSource;
        dgvResults.ColumnHeaderMouseClick += dgvResults_ColumnHeaderMouseClick;

        chkUseDateRange.Checked = _settings.UseDateRange;
        dtpFrom.Value = _settings.DateFrom;
        dtpTo.Value = _settings.DateTo;

        chkSearchBody.Checked = _settings.SearchBody;
        chkSearchAttachments.Checked = _settings.SearchAttachments;
        nudMainMaxResults.Value = Math.Clamp(_settings.MaxResults, 10, 5000);
        UpdateExcludedFoldersSummary();

        chkSearchBody.CheckedChanged += (_, _) => { _settings.SearchBody = chkSearchBody.Checked; };
        chkSearchAttachments.CheckedChanged += (_, _) => { _settings.SearchAttachments = chkSearchAttachments.Checked; };
        nudMainMaxResults.ValueChanged += (_, _) => { _settings.MaxResults = (int)nudMainMaxResults.Value; };

    RefreshHistoryDropdown();

        _columnFilterMap = new Dictionary<string, TextBox>
        {
            [nameof(EmailSearchResult.ReceivedTime)] = txtColDate,
            [nameof(EmailSearchResult.Subject)]      = txtColSubject,
        };
        _dropdownButtonMap = new Dictionary<string, Button>
        {
            [nameof(EmailSearchResult.Mailbox)]        = btnColMailbox,
            [nameof(EmailSearchResult.FolderPath)]     = btnColFolderPath,
            [nameof(EmailSearchResult.Sender)]         = btnColSender,
            [nameof(EmailSearchResult.Recipients)]     = btnColRecipients,
            [nameof(EmailSearchResult.HasAttachments)] = btnColHasAttachment,
        };
        _dropdownFilterMap = new Dictionary<string, HashSet<string>>
        {
            [nameof(EmailSearchResult.Mailbox)]        = new(StringComparer.OrdinalIgnoreCase),
            [nameof(EmailSearchResult.FolderPath)]     = new(StringComparer.OrdinalIgnoreCase),
            [nameof(EmailSearchResult.Sender)]         = new(StringComparer.OrdinalIgnoreCase),
            [nameof(EmailSearchResult.Recipients)]     = new(StringComparer.OrdinalIgnoreCase),
            [nameof(EmailSearchResult.HasAttachments)] = new(StringComparer.OrdinalIgnoreCase),
        };
        foreach (var tb in _columnFilterMap.Values)
            tb.TextChanged += (_, _) => RestartFilterDebounce();
        foreach (var (prop, btn) in _dropdownButtonMap)
        {
            string capturedProp = prop;
            Button capturedBtn = btn;
            btn.Click += (_, _) => ShowDropdownFilter(capturedBtn, capturedProp);
        }
        _filterDebounceTimer.Tick += (_, _) => { _filterDebounceTimer.Stop(); ApplyResultFilter(); };

        // Button labels depend on _dropdownButtonMap being initialized, so update them now
        UpdateAllDropdownButtonTexts();

        dgvResults.ColumnWidthChanged        += (_, _) => SyncFilterPositions();
        dgvResults.ColumnDisplayIndexChanged += (_, _) => { SyncFilterPositions(); SaveColumnOrder(); };
        dgvResults.Scroll                    += (_, _) => SyncFilterPositions();
        splitContainer.SplitterMoved         += (_, _) => SyncFilterPositions();
        Resize                               += (_, _) => SyncFilterPositions();

        chkShowPreview.CheckedChanged += (_, _) => UpdatePreviewVisibility();
        dgvResults.SelectionChanged += dgvResults_SelectionChanged;

        _indexRefreshTimer.Interval = 30000;
        _indexRefreshTimer.Tick += async (_, _) => await TryRunAutoIndexRefreshAsync();
        _indexRefreshTimer.Start();

        UpdatePreviewVisibility();
        btnCopyPreview.Enabled = false;
    }

    private async void Form1_Load(object sender, EventArgs e)
    {
        RestoreWindowBounds();
        RestoreColumnOrder();
        await RefreshStoresAsync();
        SyncFilterPositions();
        if (_settings.IndexOnStartup)
            _ = TryRunAutoIndexRefreshAsync(force: true);
    }

    private void RestoreWindowBounds()
    {
        if (_settings.WindowWidth > 0 && _settings.WindowHeight > 0)
        {
            var bounds = new Rectangle(_settings.WindowLeft, _settings.WindowTop, _settings.WindowWidth, _settings.WindowHeight);
            bool visible = Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(bounds));
            if (visible)
            {
                StartPosition = FormStartPosition.Manual;
                Location = new Point(_settings.WindowLeft, _settings.WindowTop);
                Size = new Size(_settings.WindowWidth, _settings.WindowHeight);
            }
        }

        if (_settings.WindowState == FormWindowState.Maximized)
        {
            WindowState = FormWindowState.Maximized;
        }
    }

    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
        _previewCancellation?.Cancel();
        _previewCancellation?.Dispose();
        _previewCancellation = null;
        _filterDebounceTimer.Stop();
        _filterDebounceTimer.Dispose();
        SaveWindowBounds();
        SaveSettings();
    }

    private void SaveWindowBounds()
    {
        _settings.WindowState = WindowState;
        if (WindowState == FormWindowState.Normal)
        {
            _settings.WindowLeft = Location.X;
            _settings.WindowTop = Location.Y;
            _settings.WindowWidth = Size.Width;
            _settings.WindowHeight = Size.Height;
        }
    }

    private void RestartFilterDebounce()
    {
        _filterDebounceTimer.Stop();
        _filterDebounceTimer.Start();
    }

    private void UpdateExcludedFoldersSummary()
    {
        lblExcludedFolderSummary.Text = string.Format(Strings.SettingsLblExcludedFoldersFmt, _settings.ExcludedFolderEntryIds.Count);
    }

    private async void btnExcludeFolders_Click(object sender, EventArgs e)
    {
        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            MessageBox.Show(this, Strings.MsgSelectMailbox, Strings.MsgSelectMailboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedStoreIds = selectedStores.Select(s => s.StoreId).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        bool sameStoreSelection = selectedStoreIds.Count == _folderTreeStoreIds.Count &&
            selectedStoreIds.All(id => _folderTreeStoreIds.Contains(id, StringComparer.OrdinalIgnoreCase));
        var dialogRoots = sameStoreSelection ? _folderRoots : new List<MailboxFolderRoot>();

        using var dialog = new FolderSelectionForm(
            Strings.FolderExcludeTitle,
            dialogRoots,
            _settings.ExcludedFolderEntryIds,
            async cancellationToken =>
            {
                SetBusy(true, Strings.SettingsMsgLoadingFolders);
                var roots = await OutlookSearcher.GetFolderTreeAsync(selectedStoreIds, cancellationToken);
                _folderRoots = roots.ToList();
                _folderTreeStoreIds = selectedStoreIds;
                SetBusy(false, Strings.StatusReady);
                return roots;
            });

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _settings.ExcludedFolderEntryIds = dialog.GetSelectedEntryIds().ToList();
            UpdateExcludedFoldersSummary();
            AppSettingsStore.Save(_settings);
        }
    }



    private async void btnSearch_Click(object sender, EventArgs e)
    {
        string query = cmbQuery.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            MessageBox.Show(Strings.MsgEnterQuery, Strings.MsgEnterQueryTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            cmbQuery.Focus();
            return;
        }

        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            MessageBox.Show(Strings.MsgSelectMailbox, Strings.MsgSelectMailboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var criteria = new SearchCriteria
        {
            Query = query,
            SearchBody = _settings.SearchBody,
            SearchAttachments = _settings.SearchAttachments,
            UsePersistentIndex = _settings.UsePersistentIndexForSearch,
            From = chkUseDateRange.Checked ? dtpFrom.Value.Date : null,
            To = chkUseDateRange.Checked ? dtpTo.Value.Date.AddDays(1).AddTicks(-1) : null,
            MaxResults = _settings.MaxResults,
            SelectedStoreIds = selectedStores.Select(s => s.StoreId).ToArray(),
            ExcludedFolderEntryIds = _settings.ExcludedFolderEntryIds,
            ExcludedAttachmentExtensions = _settings.ExcludeAttachmentExtensions
                ? AppSettingsParser.ParseExtensions(_settings.ExcludedAttachmentExtensionsRaw)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        };

        SetBusy(true, Strings.StatusSearchStarted);
        AddToSearchHistory(query);
        _searchCancellation = new CancellationTokenSource();
        _allResults.Clear();
        _visibleResults.Clear();
        ClearDropdownFilters();

        try
        {
            var progress = new Progress<string>(text => toolStripStatusLabel1.Text = text);
            var resultProgress = new Progress<EmailSearchResult>(result =>
            {
                _allResults.Add(result);
                if (MatchesResultFilter(result))
                {
                    _visibleResults.Add(result);
                }
            });
            var results = await OutlookSearcher.SearchAsync(criteria, progress, resultProgress, _searchCancellation.Token);

            // resultProgress callbacks are posted async to the UI message queue.
            // BeginInvoke ensures ApplyResultFilter runs after all pending callbacks have drained,
            // so _allResults is fully populated and the sort is correctly applied.
            BeginInvoke(ApplyResultFilter);

            toolStripStatusLabel1.Text = string.Format(Strings.StatusSearchDoneFmt, results.Count);
            SaveSettings();
        }
        catch (OperationCanceledException)
        {
            toolStripStatusLabel1.Text = Strings.StatusSearchCancelled;
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(
                this,
                Strings.ErrSearchTitle,
                Strings.ErrSearchSummary,
                ex,
                Strings.ErrSearchHint);
            toolStripStatusLabel1.Text = Strings.StatusSearchFailed;
        }
        finally
        {
            _searchCancellation?.Dispose();
            _searchCancellation = null;
            _previewCancellation?.Cancel();
            SetBusy(false, toolStripStatusLabel1.Text ?? string.Empty);
        }
    }

    private void btnCancel_Click(object sender, EventArgs e)
    {
        _searchCancellation?.Cancel();
    }

    private void btnCopyPreview_Click(object sender, EventArgs e)
    {
        string preview = txtPreview.Text.Trim();
        if (string.IsNullOrWhiteSpace(preview))
        {
            return;
        }

        try
        {
            Clipboard.SetText(preview);
            toolStripStatusLabel1.Text = Strings.StatusCopied;
        }
        catch (Exception ex)
        {
            toolStripStatusLabel1.Text = Strings.StatusCopyFailed;
            ErrorDetailsDialog.Show(this, Strings.ErrCopyTitle, Strings.ErrCopySummary, ex);
        }
    }

    private async void dgvResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0)
        {
            return;
        }

        if (dgvResults.Rows[e.RowIndex].DataBoundItem is not EmailSearchResult result)
        {
            return;
        }

        try
        {
            await OutlookSearcher.OpenMailItemAsync(result.EntryId, result.StoreId);
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(
                this,
                Strings.ErrOpenMailTitle,
                Strings.ErrOpenMailSummary,
                ex,
                Strings.ErrOpenMailHint);
        }
    }

    private async void dgvResults_SelectionChanged(object? sender, EventArgs e)
    {
        await RefreshPreviewAsync();
    }

    private void mnuInstellingen_Click(object sender, EventArgs e)
    {
        using var dlg = new SettingsForm(_settings, _availableStores, _folderRoots, _folderTreeStoreIds);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _availableStores = dlg.AvailableStores;
            _folderRoots = dlg.FolderRoots;
            _folderTreeStoreIds = dlg.FolderTreeStoreIds;
            // Sync controls that may have changed in settings dialog
            chkSearchBody.Checked = _settings.SearchBody;
            chkSearchAttachments.Checked = _settings.SearchAttachments;
            nudMainMaxResults.Value = Math.Clamp(_settings.MaxResults, 10, 5000);
            UpdateExcludedFoldersSummary();
            AppSettingsStore.Save(_settings);
                        RefreshHistoryDropdown();
            ApplyStrings();
        }
    }

    private void mnuHelpHelp_Click(object sender, EventArgs e)
    {
        MessageBox.Show(
            "Outlook Classic Search – Helpoverzicht\n\n" +
            "Zoeken\n" +
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
            "  • Gebruik 'Kopieer voorbeeld' om de berichttekst naar het klembord te kopiëren.",
            "Help",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void mnuHelpInfo_Click(object sender, EventArgs e)
    {
        using var info = new InfoForm();
        info.ShowDialog(this);
    }

    private async Task RefreshStoresAsync()
    {
        SetBusy(true, Strings.SettingsMsgLoadingMailboxes);
        try
        {
            var stores = await OutlookSearcher.GetAvailableStoresAsync(CancellationToken.None);
            _availableStores = stores.Select(s => new StoreListItem(s.DisplayName, s.StoreId)).ToList();
            toolStripStatusLabel1.Text = string.Format(Strings.StatusMailboxesFoundFmt, _availableStores.Count);
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(
                this,
                Strings.ErrMailboxTitle,
                Strings.ErrMailboxSummary,
                ex,
                Strings.ErrMailboxHint);
            toolStripStatusLabel1.Text = Strings.StatusMailboxesFailed;
        }
        finally
        {
            SetBusy(false, toolStripStatusLabel1.Text ?? string.Empty);
        }
    }

    private List<StoreListItem> GetSelectedStores()
    {
        if (_availableStores.Count == 0) return new List<StoreListItem>();
        var selectedIds = new HashSet<string>(_settings.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
        var selectedNames = new HashSet<string>(_settings.SelectedStoreDisplayNames, StringComparer.OrdinalIgnoreCase);
        if (selectedIds.Count == 0 && selectedNames.Count == 0) return _availableStores.ToList();
        return _availableStores
            .Where(s => selectedIds.Contains(s.StoreId) || selectedNames.Contains(s.DisplayName))
            .ToList();
    }

    private void SetBusy(bool busy, string statusText)
    {
        btnSearch.Enabled = !busy;
        btnCancel.Enabled = busy;
        toolStripProgressBar.Visible = busy;
        toolStripStatusLabel1.Text = statusText;
    }

    private void SaveSettings()
    {
        _settings.UseDateRange = chkUseDateRange.Checked;
        _settings.DateFrom = dtpFrom.Value.Date;
        _settings.DateTo = dtpTo.Value.Date;
        _settings.SearchBody = chkSearchBody.Checked;
        _settings.SearchAttachments = chkSearchAttachments.Checked;
        _settings.MaxResults = (int)nudMainMaxResults.Value;
        SaveColumnOrder();
        AppSettingsStore.Save(_settings);
    }

    private void UpdatePreviewVisibility()
    {
        splitContainer.Panel2Collapsed = !chkShowPreview.Checked;

        if (!chkShowPreview.Checked)
        {
            txtPreview.Clear();
            btnCopyPreview.Enabled = false;
            _previewCancellation?.Cancel();
            _previewCancellation?.Dispose();
            _previewCancellation = null;
        }

        _ = RefreshPreviewAsync();
    }

    private void btnClosePreview_Click(object sender, EventArgs e)
    {
        chkShowPreview.Checked = false;
    }

    private void txtQuery_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && btnSearch.Enabled)
        {
            e.SuppressKeyPress = true;
            btnSearch_Click(sender, e);
        }
    }


    private void cmbQuery_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && btnSearch.Enabled)
        {
            e.SuppressKeyPress = true;
            btnSearch_Click(sender, e);
        }
    }

    private void AddToSearchHistory(string query)
    {
        var history = _settings.SearchHistory;
        history.Remove(query);
        history.Insert(0, query);
        int max = Math.Max(1, _settings.SearchHistoryMaxCount);
        while (history.Count > max)
            history.RemoveAt(history.Count - 1);
        RefreshHistoryDropdown();
        AppSettingsStore.Save(_settings);
    }

    private void RefreshHistoryDropdown()
    {
        string current = cmbQuery.Text;
        cmbQuery.Items.Clear();
        cmbQuery.Items.AddRange(_settings.SearchHistory.ToArray());
        cmbQuery.Text = current;
    }

    private async Task TryRunAutoIndexRefreshAsync(bool force = false)
    {
        if (_isAutoIndexRunning || (!force && !_settings.IndexAutoRefreshEnabled))
        {
            return;
        }

        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            return;
        }

        var current = PersistentIndexStore.Load();
        int minutes = Math.Clamp(_settings.IndexRefreshIntervalMinutes, 5, 1440);
        if (!force && current is not null && current.BuiltAtUtc > DateTime.UtcNow.AddMinutes(-minutes))
        {
            return;
        }

        _isAutoIndexRunning = true;

        try
        {
            var request = new IndexBuildRequest
            {
                StoreIds = selectedStores.Select(s => s.StoreId).ToArray(),
                IncludedFolderEntryIds = _settings.IndexIncludedFolderEntryIds,
                SearchBody = _settings.SearchBody,
                IncludeAttachments = _settings.SearchAttachments,
                ExcludedAttachmentExtensions = AppSettingsParser.ParseExtensions(_settings.ExcludedAttachmentExtensionsRaw)
            };

            var progress = new Progress<string>(text => toolStripStatusLabel1.Text = Strings.StatusAutoIndexPrefix + text);
            await OutlookSearcher.BuildPersistentIndexAsync(request, progress, CancellationToken.None);
            toolStripStatusLabel1.Text = Strings.StatusAutoIndexDone;
        }
        catch (Exception ex)
        {
            toolStripStatusLabel1.Text = $"{Strings.StatusAutoIndexFailed} {ex.Message}";
        }
        finally
        {
            _isAutoIndexRunning = false;
        }
    }

    private void InitializeResultsGridColumns()
    {
        dgvResults.AutoGenerateColumns = false;
        dgvResults.Columns.Clear();

        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Mailbox),
            HeaderText = Strings.ColMailbox,
            Width = 150,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.FolderPath),
            HeaderText = Strings.ColFolder,
            Width = 220,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.ReceivedTime),
            HeaderText = Strings.ColDate,
            Width = 125,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy-MM-dd HH:mm" }
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Subject),
            HeaderText = Strings.ColSubject,
            Width = 260,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Sender),
            HeaderText = Strings.ColSender,
            Width = 180,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Recipients),
            HeaderText = Strings.ColRecipients,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });
        dgvResults.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.HasAttachments),
            HeaderText = Strings.ColHasAttachment,
            Width = 60,
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            FalseValue = false,
            TrueValue = true,
            ThreeState = false
        });

        UpdateSortGlyphs();
    }

    private void ApplyStrings()
    {
        // Menu
        mnuInstellingen.Text = Strings.MenuSettings;
        mnuHelp.Text         = Strings.MenuHelp;
        mnuHelpHelp.Text     = Strings.MenuHelpItem;
        mnuHelpInfo.Text     = Strings.MenuInfoItem;

        // Toolbar / controls
        lblQuery.Text              = Strings.LabelQuery;
        btnSearch.Text             = Strings.BtnSearch;
        btnCancel.Text             = Strings.BtnCancel;
        chkShowPreview.Text        = Strings.ChkShowPreview;
        chkUseDateRange.Text       = Strings.ChkUseDateRange;
        chkSearchBody.Text         = Strings.SettingsChkSearchBody;
        chkSearchAttachments.Text  = Strings.SettingsChkSearchAtt;
        lblMainMaxResults.Text     = Strings.SettingsLblMaxResults;
        btnExcludeFolders.Text     = Strings.SettingsBtnExcludeFolders;
        UpdateExcludedFoldersSummary();

        // Filter placeholders (TextBox filters)
        txtColDate.PlaceholderText    = Strings.FilterDate;
        txtColSubject.PlaceholderText = Strings.FilterSubject;

        // Dropdown filter button labels
        UpdateAllDropdownButtonTexts();

        // Preview panel
        btnClosePreview.Text = Strings.BtnClosePreview;
        btnCopyPreview.Text  = Strings.BtnCopyPreview;

        // Status bar (only reset when idle)
        if (!toolStripProgressBar.Visible)
            toolStripStatusLabel1.Text = Strings.StatusReady;

        // Column headers
        foreach (DataGridViewColumn col in dgvResults.Columns)
        {
            col.HeaderText = col.DataPropertyName switch
            {
                nameof(EmailSearchResult.Mailbox)        => Strings.ColMailbox,
                nameof(EmailSearchResult.FolderPath)     => Strings.ColFolder,
                nameof(EmailSearchResult.ReceivedTime)   => Strings.ColDate,
                nameof(EmailSearchResult.Subject)        => Strings.ColSubject,
                nameof(EmailSearchResult.Sender)         => Strings.ColSender,
                nameof(EmailSearchResult.Recipients)     => Strings.ColRecipients,
                nameof(EmailSearchResult.HasAttachments) => Strings.ColHasAttachment,
                _ => col.HeaderText
            };
        }

        SyncFilterPositions();
    }

    private void ApplyResultFilter()
    {
        _visibleResults.RaiseListChangedEvents = false;
        _visibleResults.Clear();

        foreach (var result in GetFilteredAndSortedResults())
        {
            _visibleResults.Add(result);
        }

        _visibleResults.RaiseListChangedEvents = true;
        _visibleResults.ResetBindings();
        UpdateSortGlyphs();
        _ = RefreshPreviewAsync();
    }

    private IEnumerable<EmailSearchResult> GetFilteredAndSortedResults()
    {
        var filtered = _allResults.Where(MatchesResultFilter);

        Func<EmailSearchResult, object> keySelector = _sortColumn switch
        {
            nameof(EmailSearchResult.Mailbox)        => r => r.Mailbox,
            nameof(EmailSearchResult.FolderPath)     => r => r.FolderPath,
            nameof(EmailSearchResult.Subject)        => r => r.Subject,
            nameof(EmailSearchResult.Sender)         => r => r.Sender,
            nameof(EmailSearchResult.Recipients)     => r => r.Recipients,
            nameof(EmailSearchResult.HasAttachments) => r => r.HasAttachments,
            _ => r => r.ReceivedTime
        };

        return _sortAscending
            ? filtered.OrderBy(keySelector)
            : filtered.OrderByDescending(keySelector);
    }

    private void dgvResults_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.ColumnIndex < 0 || e.ColumnIndex >= dgvResults.Columns.Count)
            return;

        if (e.Button == MouseButtons.Right)
        {
            ShowColumnVisibilityMenu();
            return;
        }

        var column = dgvResults.Columns[e.ColumnIndex];
        string? clickedProperty = column.DataPropertyName;
        if (string.IsNullOrWhiteSpace(clickedProperty))
            return;

        if (string.Equals(_sortColumn, clickedProperty, StringComparison.Ordinal))
            _sortAscending = !_sortAscending;
        else
        {
            _sortColumn = clickedProperty;
            _sortAscending = true;
        }

        ApplyResultFilter();
    }

    private void ShowColumnVisibilityMenu()
    {
        var menu = new ContextMenuStrip();
        foreach (DataGridViewColumn col in dgvResults.Columns)
        {
            var item = new ToolStripMenuItem(col.HeaderText) { Checked = col.Visible, CheckOnClick = true };
            var colRef = col;
            item.CheckedChanged += (_, _) => { colRef.Visible = item.Checked; SyncFilterPositions(); };
            menu.Items.Add(item);
        }
        menu.Show(dgvResults, dgvResults.PointToClient(Cursor.Position));
    }

    private void UpdateSortGlyphs()
    {
        foreach (DataGridViewColumn column in dgvResults.Columns)
        {
            if (string.Equals(column.DataPropertyName, _sortColumn, StringComparison.Ordinal))
            {
                column.HeaderCell.SortGlyphDirection = _sortAscending
                    ? SortOrder.Ascending
                    : SortOrder.Descending;
            }
            else
            {
                column.HeaderCell.SortGlyphDirection = SortOrder.None;
            }
        }
    }

    private bool MatchesResultFilter(EmailSearchResult result)
    {
        var comparison = StringComparison.OrdinalIgnoreCase;

        // TextBox filters (Datum, Onderwerp)
        foreach (var (prop, tb) in _columnFilterMap)
        {
            string f = tb.Text.Trim();
            if (string.IsNullOrEmpty(f)) continue;
            string value = prop switch
            {
                nameof(EmailSearchResult.ReceivedTime) => result.ReceivedTime.ToString("yyyy-MM-dd HH:mm"),
                nameof(EmailSearchResult.Subject)      => result.Subject,
                _ => string.Empty
            };
            if (!value.Contains(f, comparison)) return false;
        }

        // Dropdown multi-select filters
        foreach (var (prop, selected) in _dropdownFilterMap)
        {
            if (selected.Count == 0) continue;

            if (prop == nameof(EmailSearchResult.HasAttachments))
            {
                // Selected values are localised "Ja"/"Nee" strings
                bool wantYes = selected.Contains(Strings.FilterYes);
                bool wantNo  = selected.Contains(Strings.FilterNo);
                if (wantYes && !wantNo  && !result.HasAttachments) return false;
                if (!wantYes && wantNo  &&  result.HasAttachments) return false;
                // both selected = no filter
                continue;
            }

            string value = prop switch
            {
                nameof(EmailSearchResult.Mailbox)    => result.Mailbox,
                nameof(EmailSearchResult.FolderPath) => result.FolderPath,
                nameof(EmailSearchResult.Sender)     => result.Sender,
                nameof(EmailSearchResult.Recipients) => result.Recipients,
                _ => string.Empty
            };
            if (!selected.Contains(value, StringComparer.OrdinalIgnoreCase)) return false;
        }

        return true;
    }

    private void SyncFilterPositions()
    {
        if (_columnFilterMap.Count == 0 && _dropdownButtonMap.Count == 0) return;
        foreach (DataGridViewColumn col in dgvResults.Columns)
        {
            var rect = dgvResults.GetColumnDisplayRectangle(col.Index, false);
            bool colVisible = col.Visible && rect.Width > 0;

            if (_columnFilterMap.TryGetValue(col.DataPropertyName, out var tb))
            {
                if (!colVisible) tb.Visible = false;
                else { tb.Visible = true; tb.SetBounds(rect.X, 0, rect.Width, pnlFilters.Height); }
            }

            if (_dropdownButtonMap.TryGetValue(col.DataPropertyName, out var btn))
            {
                if (!colVisible) btn.Visible = false;
                else { btn.Visible = true; btn.SetBounds(rect.X, 0, rect.Width, pnlFilters.Height); }
            }
        }
    }

    private void ShowDropdownFilter(Button anchor, string columnProp)
    {
        if (!_dropdownFilterMap.TryGetValue(columnProp, out var selected))
        {
            selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _dropdownFilterMap[columnProp] = selected;
        }

        List<string> values;
        if (columnProp == nameof(EmailSearchResult.HasAttachments))
        {
            values = new List<string> { Strings.FilterYes, Strings.FilterNo };
        }
        else
        {
            Func<EmailSearchResult, string> getValue = columnProp switch
            {
                nameof(EmailSearchResult.Mailbox)    => r => r.Mailbox,
                nameof(EmailSearchResult.FolderPath) => r => r.FolderPath,
                nameof(EmailSearchResult.Sender)     => r => r.Sender,
                nameof(EmailSearchResult.Recipients) => r => r.Recipients,
                _ => r => string.Empty
            };

            values = _allResults
                .Select(getValue)
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToList();
        }

        var clb = new CheckedListBox
        {
            CheckOnClick = true,
            IntegralHeight = false,
            BackColor = AppTheme.Surface,
            ForeColor = AppTheme.TextPrimary,
            BorderStyle = BorderStyle.None,
            HorizontalScrollbar = true,
        };
        clb.Items.AddRange(values.ToArray<object>());

        // Dynamic size: measure widest item text, apply min/max
        int itemH = clb.ItemHeight > 0 ? clb.ItemHeight : 18;
        int measuredW = MeasureListWidth(clb, values);
        int popupW = Math.Clamp(measuredW, 120, 480);
        int popupH = Math.Clamp(values.Count * itemH + 4, 60, 320);
        clb.Size = new Size(popupW, popupH);

        // Apply checked state
        for (int i = 0; i < clb.Items.Count; i++)
        {
            if (selected.Contains((string)clb.Items[i]))
                clb.SetItemChecked(i, true);
        }

        var host = new ToolStripControlHost(clb)
        {
            Padding = new Padding(0),
            Margin = new Padding(0),
            AutoSize = false,
            Size = clb.Size
        };
        var popup = new ToolStripDropDown
        {
            AutoSize = false,
            Padding = new Padding(1),
        };
        popup.Items.Add(host);
        popup.Width  = clb.Width + 2;
        popup.Height = clb.Height + 2;

        clb.ItemCheck += (_, e) =>
        {
            string item = (string)clb.Items[e.Index];
            if (e.NewValue == CheckState.Checked) selected.Add(item);
            else selected.Remove(item);
            UpdateDropdownButtonText(columnProp, anchor);
            RestartFilterDebounce();
        };

        popup.Show(anchor, new Point(0, anchor.Height));
    }

    private static int MeasureListWidth(CheckedListBox clb, IList<string> values)
    {
        // Checkbox width (~20px) + text + padding
        const int checkboxW = 22;
        const int paddingW  = 12;
        if (values.Count == 0) return 120;
        using var g = clb.CreateGraphics();
        float maxW = values.Max(v => g.MeasureString(v, clb.Font).Width);
        return (int)Math.Ceiling(maxW) + checkboxW + paddingW;
    }

    private void UpdateDropdownButtonText(string columnProp, Button btn)
    {
        string label = columnProp switch
        {
            nameof(EmailSearchResult.Mailbox)        => Strings.ColMailbox,
            nameof(EmailSearchResult.FolderPath)     => Strings.ColFolder,
            nameof(EmailSearchResult.Sender)         => Strings.ColSender,
            nameof(EmailSearchResult.Recipients)     => Strings.ColRecipients,
            nameof(EmailSearchResult.HasAttachments) => Strings.ColHasAttachment,
            _ => columnProp
        };
        int count = _dropdownFilterMap.TryGetValue(columnProp, out var sel) ? sel.Count : 0;
        btn.Text = count > 0 ? $"{label} ({count}) \u25be" : $"{label} \u25be";
    }

    private void UpdateAllDropdownButtonTexts()
    {
        foreach (var (prop, btn) in _dropdownButtonMap)
            UpdateDropdownButtonText(prop, btn);
    }

    private void ClearDropdownFilters()
    {
        foreach (var sel in _dropdownFilterMap.Values)
            sel.Clear();
        UpdateAllDropdownButtonTexts();
    }

    private void SaveColumnOrder()
    {
        _settings.ColumnOrder = dgvResults.Columns
            .Cast<DataGridViewColumn>()
            .OrderBy(c => c.DisplayIndex)
            .Select(c => c.DataPropertyName)
            .ToList();
    }

    private void RestoreColumnOrder()
    {
        var order = _settings.ColumnOrder;
        if (order.Count == 0) return;
        for (int i = 0; i < order.Count; i++)
        {
            var col = dgvResults.Columns
                .Cast<DataGridViewColumn>()
                .FirstOrDefault(c => c.DataPropertyName == order[i]);
            if (col is not null && col.DisplayIndex != i)
                col.DisplayIndex = i;
        }
    }

    private async Task RefreshPreviewAsync()
    {
        if (!chkShowPreview.Checked)
        {
            return;
        }

        if (dgvResults.CurrentRow?.DataBoundItem is not EmailSearchResult selected)
        {
            txtPreview.Text = "Geen bericht geselecteerd.";
            btnCopyPreview.Enabled = false;
            return;
        }

        _previewCancellation?.Cancel();
        _previewCancellation?.Dispose();
        _previewCancellation = new CancellationTokenSource();

        CancellationToken token = _previewCancellation.Token;
        txtPreview.Text = "Voorbeeld laden...";
        btnCopyPreview.Enabled = false;

        try
        {
            var preview = await OutlookSearcher.GetMailPreviewAsync(selected.EntryId, selected.StoreId, token);
            if (preview is null)
            {
                txtPreview.Text = "Geen voorbeeld beschikbaar.";
                btnCopyPreview.Enabled = false;
                return;
            }

            txtPreview.Text =
                $"Onderwerp: {preview.Subject}{Environment.NewLine}" +
                $"Afzender: {preview.Sender}{Environment.NewLine}" +
                $"Geadresseerde: {preview.Recipients}{Environment.NewLine}" +
                $"Datum: {preview.ReceivedTime:yyyy-MM-dd HH:mm}{Environment.NewLine}{Environment.NewLine}" +
                preview.Body;
            txtPreview.SelectionStart = 0;
            txtPreview.SelectionLength = 0;
            btnCopyPreview.Enabled = !string.IsNullOrWhiteSpace(preview.Body) ||
                !string.IsNullOrWhiteSpace(preview.Subject) ||
                !string.IsNullOrWhiteSpace(preview.Sender) ||
                !string.IsNullOrWhiteSpace(preview.Recipients);
        }
        catch (OperationCanceledException)
        {
            // Ignore canceled preview refresh.
        }
        catch (Exception ex)
        {
            txtPreview.Text = "Voorbeeld laden mislukt: " + ex.Message;
            btnCopyPreview.Enabled = false;
        }
    }
}

internal sealed record StoreListItem(string DisplayName, string StoreId)
{
    public override string ToString() => DisplayName;
}
