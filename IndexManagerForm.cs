namespace OutlookClassicSearch;

internal sealed class IndexManagerForm : Form
{
    private readonly IReadOnlyList<StoreListItem> _selectedStores;
    private readonly AppSettings _settings;

    private readonly Label _lblScope = new();
    private readonly Label _lblIndexState = new();
    private readonly Label _lblInterval = new();
    private readonly CheckBox _chkAutoRefresh = new();
    private readonly CheckBox _chkIndexOnStartup = new();
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

        Text = Strings.IndexTitle;
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 380);
        MinimumSize = new Size(640, 300);
        AppVisualAssets.ApplyWindowIcon(this);

        _lblScope.AutoSize = false;
        _lblScope.Dock = DockStyle.Top;
        _lblScope.Height = 44;
        _lblScope.Padding = new Padding(10, 10, 10, 0);
        _lblScope.Text = Strings.IndexScopeNotLoaded;

        _lblIndexState.AutoSize = false;
        _lblIndexState.Dock = DockStyle.Top;
        _lblIndexState.Height = 44;
        _lblIndexState.Padding = new Padding(10, 0, 10, 0);

        var panel = new Panel { Dock = DockStyle.Fill };

        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 48 };
        _btnClose.Text = Strings.BtnClose;
        _btnClose.SetBounds(0, 8, 100, 32);
        _btnClose.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        _btnClose.Click += (_, _) => Close();
        bottomPanel.Controls.Add(_btnClose);
        bottomPanel.Layout += (_, _) =>
        {
            _btnClose.Left = bottomPanel.Width - _btnClose.Width - 12;
        };

        _btnSelectFolders.Text = Strings.IndexBtnSelectFolders;
        _btnSelectFolders.SetBounds(12, 12, 220, 32);
        _btnSelectFolders.Click += async (_, _) => await SelectIncludedFoldersAsync();

        _btnBuild.Text = Strings.IndexBtnRefresh;
        _btnBuild.SetBounds(240, 12, 180, 32);
        _btnBuild.Click += async (_, _) => await BuildIndexAsync();

        _btnClear.Text = Strings.IndexBtnClear;
        _btnClear.SetBounds(428, 12, 150, 32);
        _btnClear.Click += (_, _) =>
        {
            PersistentIndexStore.Clear();
            _lblIndexState.Text = Strings.IndexStateNone;
        };

        _chkAutoRefresh.Text = Strings.IndexChkAutoRefresh;
        _chkAutoRefresh.SetBounds(12, 66, 170, 24);
        _chkAutoRefresh.Checked = _settings.IndexAutoRefreshEnabled;
        _chkAutoRefresh.CheckedChanged += (_, _) => _settings.IndexAutoRefreshEnabled = _chkAutoRefresh.Checked;

        _lblInterval.Text = Strings.IndexLblInterval;
        _lblInterval.SetBounds(200, 68, 90, 20);

        _nudInterval.Minimum = 5;
        _nudInterval.Maximum = 1440;
        _nudInterval.Value = Math.Clamp(_settings.IndexRefreshIntervalMinutes, 5, 1440);
        _nudInterval.SetBounds(290, 66, 80, 24);
        _nudInterval.ValueChanged += (_, _) => _settings.IndexRefreshIntervalMinutes = (int)_nudInterval.Value;

        _chkIndexOnStartup.Text = Strings.IndexChkIndexOnStartup;
        _chkIndexOnStartup.SetBounds(390, 66, 220, 24);
        _chkIndexOnStartup.Checked = _settings.IndexOnStartup;
        _chkIndexOnStartup.CheckedChanged += (_, _) => _settings.IndexOnStartup = _chkIndexOnStartup.Checked;

        _progress.Style = ProgressBarStyle.Marquee;
        _progress.Visible = false;
        _progress.SetBounds(12, 110, 720, 22);
        _progress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

        panel.Controls.Add(_btnSelectFolders);
        panel.Controls.Add(_btnBuild);
        panel.Controls.Add(_btnClear);
        panel.Controls.Add(_chkAutoRefresh);
        panel.Controls.Add(_lblInterval);
        panel.Controls.Add(_nudInterval);
        panel.Controls.Add(_chkIndexOnStartup);
        panel.Controls.Add(_progress);

        Controls.Add(panel);
        Controls.Add(bottomPanel);
        Controls.Add(_lblIndexState);
        Controls.Add(_lblScope);

        AppTheme.Apply(this);
        AppTheme.ApplyPrimaryStyle(_btnBuild);

        Shown += (_, _) => LoadData();
    }

    private void LoadData()
    {
        _lblScope.Text = string.Format(Strings.IndexScopeSelectedFmt, _selectedStores.Count);

        var existing = PersistentIndexStore.Load();
        _lblIndexState.Text = existing is null
            ? Strings.IndexStateNone
            : string.Format(Strings.IndexStateFmt, existing.Items.Count, existing.BuiltAtUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
    }

    private async Task SelectIncludedFoldersAsync()
    {
        using var dialog = new FolderSelectionForm(
            Strings.IndexFolderTitle,
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
            MessageBox.Show(this, Strings.IndexMsgNoMailbox, Strings.IndexMsgNoMailboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            var progress = new Progress<string>(text => _lblIndexState.Text = string.Format(Strings.IndexBuildProgressFmt, text));
            var built = await OutlookSearcher.BuildPersistentIndexAsync(request, progress, CancellationToken.None);
            _lblIndexState.Text = string.Format(Strings.IndexStateFmt, built.Items.Count, built.BuiltAtUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(this, Strings.IndexErrTitle, Strings.IndexErrSummary, ex);
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
        _chkIndexOnStartup.Enabled = enabled;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_isBuildingIndex)
        {
            e.Cancel = true;
            MessageBox.Show(this, Strings.IndexMsgBusy, Strings.IndexMsgBusyTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        AppSettingsStore.Save(_settings);
        base.OnFormClosing(e);
    }
}
