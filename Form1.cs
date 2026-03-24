using System.ComponentModel;

namespace OutlookClassicSearch;

public partial class Form1 : Form
{
    private const int ResultsTop = 352;
    private const int PreviewHeight = 170;
    private const int ResultsBottomMargin = 8;

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
    private string _sortColumn = nameof(EmailSearchResult.ReceivedTime);
    private bool _sortAscending;

    public Form1()
    {
        InitializeComponent();

        _settings = AppSettingsStore.Load();
        InitializeResultsGridColumns();
        _resultsBindingSource.DataSource = _visibleResults;
        dgvResults.DataSource = _resultsBindingSource;
        dgvResults.ColumnHeaderMouseClick += dgvResults_ColumnHeaderMouseClick;

        txtExcludedAttachmentExt.Text = _settings.ExcludedAttachmentExtensionsRaw;
        chkSearchBody.Checked = _settings.SearchBody;
        chkSearchAttachments.Checked = _settings.SearchAttachments;
        chkUseDateRange.Checked = _settings.UseDateRange;
        dtpFrom.Value = _settings.DateFrom;
        dtpTo.Value = _settings.DateTo;
        nudMaxResults.Value = Math.Clamp(_settings.MaxResults, 10, 5000);
        chkUseIndex.Checked = _settings.UsePersistentIndexForSearch;
        chkExcludeAttachmentExt.Checked = _settings.ExcludeAttachmentExtensions;
        txtExcludedAttachmentExt.Enabled = _settings.ExcludeAttachmentExtensions;
        chkExcludeAttachmentExt.CheckedChanged += (_, _) => txtExcludedAttachmentExt.Enabled = chkExcludeAttachmentExt.Checked;

        txtResultFilter.TextChanged += (_, _) => ApplyResultFilter();
        txtFilterMailbox.TextChanged += (_, _) => ApplyResultFilter();
        txtFilterRecipients.TextChanged += (_, _) => ApplyResultFilter();
        chkShowPreview.CheckedChanged += (_, _) => UpdatePreviewVisibility();
        dgvResults.SelectionChanged += dgvResults_SelectionChanged;
        Resize += (_, _) => UpdateResultsLayout();

        _indexRefreshTimer.Interval = 30000;
        _indexRefreshTimer.Tick += async (_, _) => await TryRunAutoIndexRefreshAsync();
        _indexRefreshTimer.Start();

        UpdateResultsLayout();
        UpdateExcludedFolderSummary();
        btnCopyPreview.Enabled = false;
    }

    private async void Form1_Load(object sender, EventArgs e)
    {
        await RefreshStoresAsync();
    }

    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
        _previewCancellation?.Cancel();
        _previewCancellation?.Dispose();
        _previewCancellation = null;
        SaveSettings();
    }

    private async void btnRefreshStores_Click(object sender, EventArgs e)
    {
        CaptureSelectedStoresToSettings();
        _folderRoots = new List<MailboxFolderRoot>();
        _folderTreeStoreIds = new List<string>();
        await RefreshStoresAsync();
    }

    private async void btnSearch_Click(object sender, EventArgs e)
    {
        string query = txtQuery.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            MessageBox.Show("Vul een zoekterm in.", "Zoekterm ontbreekt", MessageBoxButtons.OK, MessageBoxIcon.Information);
            txtQuery.Focus();
            return;
        }

        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            MessageBox.Show("Selecteer minimaal 1 mailbox.", "Mailbox ontbreekt", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var criteria = new SearchCriteria
        {
            Query = query,
            SearchBody = chkSearchBody.Checked,
            SearchAttachments = chkSearchAttachments.Checked,
            UsePersistentIndex = chkUseIndex.Checked,
            From = chkUseDateRange.Checked ? dtpFrom.Value.Date : null,
            To = chkUseDateRange.Checked ? dtpTo.Value.Date.AddDays(1).AddTicks(-1) : null,
            MaxResults = (int)nudMaxResults.Value,
            SelectedStoreIds = selectedStores.Select(s => s.StoreId).ToArray(),
            ExcludedFolderEntryIds = _settings.ExcludedFolderEntryIds,
            ExcludedAttachmentExtensions = chkExcludeAttachmentExt.Checked
                ? AppSettingsParser.ParseExtensions(txtExcludedAttachmentExt.Text)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        };

        SetBusy(true, "Zoeken gestart...");
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

            toolStripStatusLabel1.Text = $"Klaar. {results.Count} resultaat/resultaten.";
            SaveSettings();
        }
        catch (OperationCanceledException)
        {
            toolStripStatusLabel1.Text = "Zoeken geannuleerd.";
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(
                this,
                "Zoeken mislukt",
                "Zoeken in Outlook is mislukt.",
                ex,
                "Controleer of Outlook Classic draait en of je mailboxen bereikbaar zijn.");
            toolStripStatusLabel1.Text = "Zoeken mislukt.";
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
            toolStripStatusLabel1.Text = "Voorbeeld gekopieerd naar klembord.";
        }
        catch (Exception ex)
        {
            toolStripStatusLabel1.Text = "Kopieren mislukt.";
            ErrorDetailsDialog.Show(this, "Kopieren mislukt", "Het voorbeeld kon niet worden gekopieerd.", ex);
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
                "Mail openen mislukt",
                "De geselecteerde mail kon niet worden geopend.",
                ex,
                "De mail is mogelijk verwijderd of je hebt geen rechten meer op deze mailbox.");
        }
    }

    private async void dgvResults_SelectionChanged(object? sender, EventArgs e)
    {
        await RefreshPreviewAsync();
    }

    private async void btnChooseExcludedFolders_Click(object sender, EventArgs e)
    {
        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            MessageBox.Show(this, "Selecteer eerst minimaal 1 mailbox.", "Geen mailboxen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedStoreIds = selectedStores.Select(s => s.StoreId).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        bool sameStoreSelection = selectedStoreIds.Count == _folderTreeStoreIds.Count &&
            selectedStoreIds.All(id => _folderTreeStoreIds.Contains(id, StringComparer.OrdinalIgnoreCase));
        var dialogRoots = sameStoreSelection ? _folderRoots : new List<MailboxFolderRoot>();

        using var dialog = new FolderSelectionForm(
            "Mappen uitsluiten van zoeken",
            dialogRoots,
            _settings.ExcludedFolderEntryIds,
            async cancellationToken =>
            {
                toolStripStatusLabel1.Text = "Mappenstructuur laden...";
                var roots = await OutlookSearcher.GetFolderTreeAsync(selectedStoreIds, cancellationToken);
                _folderRoots = roots.ToList();
                _folderTreeStoreIds = selectedStoreIds;
                toolStripStatusLabel1.Text = "Mappenstructuur geladen.";
                return roots;
            });

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _settings.ExcludedFolderEntryIds = dialog.GetSelectedEntryIds().ToList();
            UpdateExcludedFolderSummary();
            SaveSettings();
        }
    }

    private void btnManageIndex_Click(object sender, EventArgs e)
    {
        var selectedStores = GetSelectedStores();
        using var form = new IndexManagerForm(selectedStores, _settings, _folderRoots);
        if (form.ShowDialog(this) == DialogResult.OK)
        {
            SaveSettings();
        }
    }

    private async Task RefreshStoresAsync()
    {
        SetBusy(true, "Mailboxen ophalen...");
        try
        {
            var stores = await OutlookSearcher.GetAvailableStoresAsync(CancellationToken.None);
            clbStores.Items.Clear();

            var selectedIds = new HashSet<string>(_settings.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
            var selectedNames = new HashSet<string>(_settings.SelectedStoreDisplayNames, StringComparer.OrdinalIgnoreCase);

            foreach (var store in stores)
            {
                bool hasSavedSelection = selectedIds.Count > 0 || selectedNames.Count > 0;
                bool isChecked = !hasSavedSelection || selectedIds.Contains(store.StoreId) || selectedNames.Contains(store.DisplayName);
                clbStores.Items.Add(new StoreListItem(store.DisplayName, store.StoreId), isChecked);
            }

            toolStripStatusLabel1.Text = $"{stores.Count} mailbox(en) gevonden.";
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(
                this,
                "Mailboxen laden mislukt",
                "Mailboxen konden niet uit Outlook worden opgehaald.",
                ex,
                "Start Outlook Classic handmatig, controleer je profiel en probeer daarna opnieuw.");
            toolStripStatusLabel1.Text = "Mailboxen laden mislukt.";
        }
        finally
        {
            SetBusy(false, toolStripStatusLabel1.Text ?? string.Empty);
        }
    }

    private async Task LoadFolderTreeAsync()
    {
        var selectedStores = GetSelectedStores();
        if (selectedStores.Count == 0)
        {
            _folderRoots = new List<MailboxFolderRoot>();
            return;
        }

        try
        {
            toolStripStatusLabel1.Text = "Mappenstructuur laden...";
            _folderRoots = (await OutlookSearcher.GetFolderTreeAsync(selectedStores.Select(s => s.StoreId).ToArray(), CancellationToken.None)).ToList();
            UpdateExcludedFolderSummary();
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(this, "Mappen laden mislukt", "De mappenstructuur kon niet worden geladen.", ex);
        }
    }

    private List<StoreListItem> GetSelectedStores()
    {
        return clbStores.CheckedItems
            .OfType<StoreListItem>()
            .ToList();
    }

    private void SetBusy(bool busy, string statusText)
    {
        btnSearch.Enabled = !busy;
        btnRefreshStores.Enabled = !busy;
        btnCancel.Enabled = busy;
        btnChooseExcludedFolders.Enabled = !busy;
        btnManageIndex.Enabled = !busy;
        chkUseIndex.Enabled = !busy;
        progressBar.Visible = busy;
        progressBar.Style = ProgressBarStyle.Marquee;
        toolStripStatusLabel1.Text = statusText;
    }

    private void SaveSettings()
    {
        CaptureSelectedStoresToSettings();
        _settings.SearchBody = chkSearchBody.Checked;
        _settings.SearchAttachments = chkSearchAttachments.Checked;
        _settings.UseDateRange = chkUseDateRange.Checked;
        _settings.DateFrom = dtpFrom.Value.Date;
        _settings.DateTo = dtpTo.Value.Date;
        _settings.MaxResults = (int)nudMaxResults.Value;
        _settings.UsePersistentIndexForSearch = chkUseIndex.Checked;
        _settings.ExcludeAttachmentExtensions = chkExcludeAttachmentExt.Checked;
        _settings.ExcludedAttachmentExtensionsRaw = txtExcludedAttachmentExt.Text.Trim();

        AppSettingsStore.Save(_settings);
    }

    private void UpdatePreviewVisibility()
    {
        UpdateResultsLayout();

        if (!chkShowPreview.Checked)
        {
            txtPreview.Clear();
            btnCopyPreview.Enabled = false;
            _previewCancellation?.Cancel();
            _previewCancellation?.Dispose();
            _previewCancellation = null;
        }

        ApplyResultFilter();
        _ = RefreshPreviewAsync();
    }

    private void UpdateResultsLayout()
    {
        int statusTop = ClientSize.Height - statusStrip1.Height;

        if (chkShowPreview.Checked)
        {
            pnlPreview.Visible = true;
            pnlPreview.Height = PreviewHeight;
            pnlPreview.Top = Math.Max(ResultsTop + 80, statusTop - ResultsBottomMargin - PreviewHeight);
            dgvResults.Top = ResultsTop;
            dgvResults.Height = Math.Max(80, pnlPreview.Top - ResultsTop - 6);
        }
        else
        {
            pnlPreview.Visible = false;
            dgvResults.Top = ResultsTop;
            dgvResults.Height = Math.Max(80, statusTop - ResultsBottomMargin - ResultsTop);
        }
    }

    private void CaptureSelectedStoresToSettings()
    {
        var selected = GetSelectedStores();
        _settings.SelectedStoreIds = selected.Select(s => s.StoreId).ToList();
        _settings.SelectedStoreDisplayNames = selected.Select(s => s.DisplayName).ToList();
    }

    private void UpdateExcludedFolderSummary()
    {
        lblExcludedFolderSummary.Text = $"Uitgesloten mappen: {_settings.ExcludedFolderEntryIds.Count}";
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

            var progress = new Progress<string>(text => toolStripStatusLabel1.Text = "Auto-index: " + text);
            await OutlookSearcher.BuildPersistentIndexAsync(request, progress, CancellationToken.None);
            toolStripStatusLabel1.Text = "Auto-index bijgewerkt.";
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
            HeaderText = "Mailbox",
            Width = 150,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.FolderPath),
            HeaderText = "Map",
            Width = 220,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.ReceivedTime),
            HeaderText = "Datum",
            Width = 125,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy-MM-dd HH:mm" }
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Subject),
            HeaderText = "Onderwerp",
            Width = 260,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Sender),
            HeaderText = "Afzender",
            Width = 180,
            SortMode = DataGridViewColumnSortMode.Programmatic
        });
        dgvResults.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(EmailSearchResult.Recipients),
            HeaderText = "Geadresseerde",
            SortMode = DataGridViewColumnSortMode.Programmatic,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });

        UpdateSortGlyphs();
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
        {
            return;
        }

        var column = dgvResults.Columns[e.ColumnIndex];
        string? clickedProperty = column.DataPropertyName;
        if (string.IsNullOrWhiteSpace(clickedProperty))
        {
            return;
        }

        if (string.Equals(_sortColumn, clickedProperty, StringComparison.Ordinal))
        {
            _sortAscending = !_sortAscending;
        }
        else
        {
            _sortColumn = clickedProperty;
            _sortAscending = true;
        }

        ApplyResultFilter();
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
        string globalFilter = txtResultFilter.Text.Trim();
        string mailboxFilter = txtFilterMailbox.Text.Trim();
        string recipientsFilter = txtFilterRecipients.Text.Trim();

        var comparison = StringComparison.OrdinalIgnoreCase;

        bool mailboxMatches = string.IsNullOrWhiteSpace(mailboxFilter) ||
            result.Mailbox.Contains(mailboxFilter, comparison);
        bool recipientsMatches = string.IsNullOrWhiteSpace(recipientsFilter) ||
            result.Recipients.Contains(recipientsFilter, comparison);
        bool globalMatches = string.IsNullOrWhiteSpace(globalFilter) ||
            result.Mailbox.Contains(globalFilter, comparison) ||
            result.FolderPath.Contains(globalFilter, comparison) ||
            result.ReceivedTime.ToString("yyyy-MM-dd HH:mm").Contains(globalFilter, comparison) ||
            result.Subject.Contains(globalFilter, comparison) ||
            result.Sender.Contains(globalFilter, comparison) ||
            result.Recipients.Contains(globalFilter, comparison);

        return mailboxMatches && recipientsMatches && globalMatches;
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
