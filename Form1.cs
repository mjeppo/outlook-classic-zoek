using System.ComponentModel;

namespace OutlookClassicSearch;

public partial class Form1 : Form
{
    private readonly BindingSource _resultsBindingSource = new();
    private readonly BindingList<EmailSearchResult> _visibleResults = new();
    private readonly List<EmailSearchResult> _allResults = new();
    private readonly AppSettings _settings;
    private readonly System.Windows.Forms.Timer _indexRefreshTimer = new();

    private CancellationTokenSource? _searchCancellation;
    private CancellationTokenSource? _previewCancellation;
    private bool _isAutoIndexRunning;
    private List<MailboxFolderRoot> _folderRoots = new();
    private List<string> _folderTreeStoreIds = new();
    private List<StoreListItem> _availableStores = new();
    private Dictionary<string, TextBox> _columnFilterMap = new();
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

    RefreshHistoryDropdown();

        _columnFilterMap = new Dictionary<string, TextBox>
        {
            [nameof(EmailSearchResult.Mailbox)]      = txtColMailbox,
            [nameof(EmailSearchResult.FolderPath)]   = txtColFolderPath,
            [nameof(EmailSearchResult.ReceivedTime)] = txtColDate,
            [nameof(EmailSearchResult.Subject)]      = txtColSubject,
            [nameof(EmailSearchResult.Sender)]       = txtColSender,
            [nameof(EmailSearchResult.Recipients)]   = txtColRecipients,
        };
        foreach (var tb in _columnFilterMap.Values)
            tb.TextChanged += (_, _) => ApplyResultFilter();

        dgvResults.ColumnWidthChanged        += (_, _) => SyncFilterPositions();
        dgvResults.ColumnDisplayIndexChanged += (_, _) => SyncFilterPositions();
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
        await RefreshStoresAsync();
        SyncFilterPositions();
    }

    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
        _previewCancellation?.Cancel();
        _previewCancellation?.Dispose();
        _previewCancellation = null;
        SaveSettings();
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

            ApplyResultFilter();

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

    private async Task TryRunAutoIndexRefreshAsync()
    {
        if (_isAutoIndexRunning || !_settings.IndexAutoRefreshEnabled)
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
        if (current is not null && current.BuiltAtUtc > DateTime.UtcNow.AddMinutes(-minutes))
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
        catch
        {
            // Silent for timer-based refresh.
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

        // Filter placeholders
        txtColMailbox.PlaceholderText    = Strings.FilterMailbox;
        txtColFolderPath.PlaceholderText = Strings.FilterFolder;
        txtColDate.PlaceholderText       = Strings.FilterDate;
        txtColSubject.PlaceholderText    = Strings.FilterSubject;
        txtColSender.PlaceholderText     = Strings.FilterSender;
        txtColRecipients.PlaceholderText = Strings.FilterRecipients;

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
                nameof(EmailSearchResult.Mailbox)      => Strings.ColMailbox,
                nameof(EmailSearchResult.FolderPath)   => Strings.ColFolder,
                nameof(EmailSearchResult.ReceivedTime) => Strings.ColDate,
                nameof(EmailSearchResult.Subject)      => Strings.ColSubject,
                nameof(EmailSearchResult.Sender)       => Strings.ColSender,
                nameof(EmailSearchResult.Recipients)   => Strings.ColRecipients,
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
            nameof(EmailSearchResult.Mailbox) => r => r.Mailbox,
            nameof(EmailSearchResult.FolderPath) => r => r.FolderPath,
            nameof(EmailSearchResult.Subject) => r => r.Subject,
            nameof(EmailSearchResult.Sender) => r => r.Sender,
            nameof(EmailSearchResult.Recipients) => r => r.Recipients,
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
        foreach (var (prop, tb) in _columnFilterMap)
        {
            string f = tb.Text.Trim();
            if (string.IsNullOrEmpty(f)) continue;
            string value = prop switch
            {
                nameof(EmailSearchResult.Mailbox)      => result.Mailbox,
                nameof(EmailSearchResult.FolderPath)   => result.FolderPath,
                nameof(EmailSearchResult.ReceivedTime) => result.ReceivedTime.ToString("yyyy-MM-dd HH:mm"),
                nameof(EmailSearchResult.Subject)      => result.Subject,
                nameof(EmailSearchResult.Sender)       => result.Sender,
                nameof(EmailSearchResult.Recipients)   => result.Recipients,
                _ => string.Empty
            };
            if (!value.Contains(f, comparison)) return false;
        }
        return true;
    }

    private void SyncFilterPositions()
    {
        if (_columnFilterMap.Count == 0) return;
        foreach (DataGridViewColumn col in dgvResults.Columns)
        {
            if (!_columnFilterMap.TryGetValue(col.DataPropertyName, out var tb)) continue;
            var rect = dgvResults.GetColumnDisplayRectangle(col.Index, false);
            if (!col.Visible || rect.Width == 0)
            {
                tb.Visible = false;
            }
            else
            {
                tb.Visible = true;
                tb.SetBounds(rect.X, 0, rect.Width, pnlFilters.Height);
            }
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
