namespace OutlookClassicSearch;

internal sealed class IndexManagerForm : Form
{
    private readonly IReadOnlyList<StoreListItem> _selectedStores;
    private readonly AppSettings _settings;

    private readonly Label _lblScope = new();
    private readonly Label _lblIndexState = new();
    private readonly Label _lblInterval = new();
    private readonly CheckBox _chkAutoRefresh = new();
    private readonly NumericUpDown _nudInterval = new();
    private readonly ProgressBar _progress = new();
    private readonly Button _btnSelectFolders = new();
    private readonly Button _btnBuild = new();
    private readonly Button _btnClear = new();
    private readonly Button _btnClose = new();

    private List<MailboxFolderRoot> _roots = new();
    private bool _isBuildingIndex;

    public IndexManagerForm(IReadOnlyList<StoreListItem> selectedStores, AppSettings settings, IReadOnlyList<MailboxFolderRoot> preloadedRoots)
    {
        _selectedStores = selectedStores;
        _settings = settings;
        _roots = new List<MailboxFolderRoot>(preloadedRoots);

        Text = "Indexbeheer";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 380);
        MinimumSize = new Size(640, 300);

        _lblScope.AutoSize = false;
        _lblScope.Dock = DockStyle.Top;
        _lblScope.Height = 44;
        _lblScope.Padding = new Padding(10, 10, 10, 0);
        _lblScope.Text = "Mailboxscope: niet geladen";

        _lblIndexState.AutoSize = false;
        _lblIndexState.Dock = DockStyle.Top;
        _lblIndexState.Height = 44;
        _lblIndexState.Padding = new Padding(10, 0, 10, 0);

        var panel = new Panel { Dock = DockStyle.Fill };

        _btnSelectFolders.Text = "Mappen voor index kiezen";
        _btnSelectFolders.SetBounds(12, 12, 220, 32);
        _btnSelectFolders.Click += async (_, _) => await SelectIncludedFoldersAsync();

        _btnBuild.Text = "Index nu verversen";
        _btnBuild.SetBounds(240, 12, 180, 32);
        _btnBuild.Click += async (_, _) => await BuildIndexAsync();

        _btnClear.Text = "Index verwijderen";
        _btnClear.SetBounds(428, 12, 150, 32);
        _btnClear.Click += (_, _) =>
        {
            PersistentIndexStore.Clear();
            _lblIndexState.Text = "Index: verwijderd";
        };

        _chkAutoRefresh.Text = "Automatisch verversen";
        _chkAutoRefresh.SetBounds(12, 66, 170, 24);
        _chkAutoRefresh.Checked = _settings.IndexAutoRefreshEnabled;
        _chkAutoRefresh.CheckedChanged += (_, _) => _settings.IndexAutoRefreshEnabled = _chkAutoRefresh.Checked;

        _lblInterval.Text = "Interval (min):";
        _lblInterval.SetBounds(200, 68, 90, 20);

        _nudInterval.Minimum = 5;
        _nudInterval.Maximum = 1440;
        _nudInterval.Value = Math.Clamp(_settings.IndexRefreshIntervalMinutes, 5, 1440);
        _nudInterval.SetBounds(290, 66, 80, 24);
        _nudInterval.ValueChanged += (_, _) => _settings.IndexRefreshIntervalMinutes = (int)_nudInterval.Value;

        _progress.Style = ProgressBarStyle.Marquee;
        _progress.Visible = false;
        _progress.SetBounds(12, 110, 720, 22);
        _progress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

        _btnClose.Text = "Sluiten";
        _btnClose.SetBounds(632, 270, 100, 32);
        _btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        _btnClose.Click += (_, _) => DialogResult = DialogResult.OK;

        panel.Controls.Add(_btnSelectFolders);
        panel.Controls.Add(_btnBuild);
        panel.Controls.Add(_btnClear);
        panel.Controls.Add(_chkAutoRefresh);
        panel.Controls.Add(_lblInterval);
        panel.Controls.Add(_nudInterval);
        panel.Controls.Add(_progress);
        panel.Controls.Add(_btnClose);

        Controls.Add(panel);
        Controls.Add(_lblIndexState);
        Controls.Add(_lblScope);

        Shown += async (_, _) => await LoadDataAsync();
    }

    private async Task LoadDataAsync()
    {
        _lblScope.Text = $"Mailboxscope: {_selectedStores.Count} geselecteerd";

        var existing = PersistentIndexStore.Load();
        _lblIndexState.Text = existing is null
            ? "Index: nog niet opgebouwd"
            : $"Index: {existing.Items.Count} items, laatst: {existing.BuiltAtUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}";

        await Task.CompletedTask;
    }

    private async Task SelectIncludedFoldersAsync()
    {
        using var dialog = new FolderSelectionForm(
            "Mappen opnemen in index",
            _roots,
            _settings.IndexIncludedFolderEntryIds,
            async cancellationToken =>
            {
                if (_selectedStores.Count == 0)
                {
                    return Array.Empty<MailboxFolderRoot>();
                }

                var roots = await OutlookSearcher.GetFolderTreeAsync(_selectedStores.Select(s => s.StoreId).ToArray(), cancellationToken);
                _roots = roots.ToList();
                return roots;
            });

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _settings.IndexIncludedFolderEntryIds = dialog.GetSelectedEntryIds().ToList();
            AppSettingsStore.Save(_settings);
        }
    }

    private async Task BuildIndexAsync()
    {
        if (_selectedStores.Count == 0)
        {
            MessageBox.Show(this, "Selecteer eerst mailboxen in het hoofdscherm.", "Geen mailboxen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _isBuildingIndex = true;
        _btnClose.Enabled = false;
        _progress.Visible = true;
        ToggleButtons(false);

        try
        {
            var request = new IndexBuildRequest
            {
                StoreIds = _selectedStores.Select(s => s.StoreId).ToArray(),
                IncludedFolderEntryIds = _settings.IndexIncludedFolderEntryIds,
                SearchBody = _settings.SearchBody,
                IncludeAttachments = _settings.SearchAttachments,
                ExcludedAttachmentExtensions = AppSettingsParser.ParseExtensions(_settings.ExcludedAttachmentExtensionsRaw)
            };

            var progress = new Progress<string>(text => _lblIndexState.Text = $"Index: {text}");
            var built = await OutlookSearcher.BuildPersistentIndexAsync(request, progress, CancellationToken.None);
            _lblIndexState.Text = $"Index: {built.Items.Count} items, laatst: {built.BuiltAtUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}";
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(this, "Indexeren mislukt", "Het opbouwen van de index is mislukt.", ex);
        }
        finally
        {
            _isBuildingIndex = false;
            _btnClose.Enabled = true;
            _progress.Visible = false;
            ToggleButtons(true);
            AppSettingsStore.Save(_settings);
        }
    }

    private void ToggleButtons(bool enabled)
    {
        _btnSelectFolders.Enabled = enabled;
        _btnBuild.Enabled = enabled;
        _btnClear.Enabled = enabled;
        _chkAutoRefresh.Enabled = enabled;
        _nudInterval.Enabled = enabled;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_isBuildingIndex)
        {
            e.Cancel = true;
            MessageBox.Show(this, "Wacht tot de indexering klaar is voordat je dit venster sluit.", "Indexering actief", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        base.OnFormClosing(e);
    }
}
