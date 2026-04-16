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
    private readonly Button _btnRefresh = new();
    private readonly Button _btnRebuild = new();
    private readonly Button _btnCancel = new();
    private readonly Button _btnClear = new();
    private readonly Button _btnClose = new();

    private List<MailboxFolderRoot> _roots = new();
    private bool _isBuildingIndex;
    private CancellationTokenSource? _buildCts;

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

        _btnRefresh.Text = Strings.IndexBtnRefresh;
        _btnRefresh.SetBounds(240, 12, 160, 32);
        _btnRefresh.Click += async (_, _) => await BuildIndexAsync(forceRebuild: false);

        _btnRebuild.Text = Strings.IndexBtnRebuild;
        _btnRebuild.SetBounds(408, 12, 180, 32);
        _btnRebuild.Click += async (_, _) => await BuildIndexAsync(forceRebuild: true);

        _btnCancel.Text = Strings.BtnCancel;
        _btnCancel.SetBounds(240, 12, 348, 32);
        _btnCancel.Visible = false;
        _btnCancel.Click += (_, _) =>
        {
            _buildCts?.Cancel();
            _btnCancel.Enabled = false;
        };

        _btnClear.Text = Strings.IndexBtnClear;
        _btnClear.SetBounds(596, 12, 130, 32);
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
        panel.Controls.Add(_btnRefresh);
        panel.Controls.Add(_btnRebuild);
        panel.Controls.Add(_btnCancel);
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
        AppTheme.ApplyPrimaryStyle(_btnRefresh);

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

    private async Task BuildIndexAsync(bool forceRebuild = false)
    {
        if (_selectedStores.Count == 0)
        {
            MessageBox.Show(this, Strings.IndexMsgNoMailbox, Strings.IndexMsgNoMailboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _isBuildingIndex = true;
        _buildCts = new CancellationTokenSource();
        _btnClose.Enabled = false;
        _progress.Visible = true;
        _btnRefresh.Visible = false;
        _btnRebuild.Visible = false;
        _btnCancel.Visible = true;
        _btnCancel.Enabled = true;
        ToggleButtons(false);

        try
        {
            var request = new IndexBuildRequest
            {
                StoreIds = _selectedStores.Select(s => s.StoreId).ToArray(),
                IncludedFolderEntryIds = _settings.IndexIncludedFolderEntryIds,
                SearchBody = _settings.SearchBody,
                IncludeAttachments = _settings.SearchAttachments,
                ExcludedAttachmentExtensions = AppSettingsParser.ParseExtensions(_settings.ExcludedAttachmentExtensionsRaw),
                ForceRebuild = forceRebuild
            };

            var progress = new Progress<string>(text => _lblIndexState.Text = string.Format(Strings.IndexBuildProgressFmt, text));
            var built = await OutlookSearcher.BuildPersistentIndexAsync(request, progress, _buildCts.Token);
            _lblIndexState.Text = string.Format(Strings.IndexStateFmt, built.Items.Count, built.BuiltAtUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
        }
        catch (OperationCanceledException)
        {
            _lblIndexState.Text = IsEnglish ? "Indexing cancelled." : "Indexeren afgebroken.";
        }
        catch (Exception ex)
        {
            ErrorDetailsDialog.Show(this, Strings.IndexErrTitle, Strings.IndexErrSummary, ex);
        }
        finally
        {
            _buildCts.Dispose();
            _buildCts = null;
            _isBuildingIndex = false;
            _btnClose.Enabled = true;
            _progress.Visible = false;
            _btnRefresh.Visible = true;
            _btnRebuild.Visible = true;
            _btnCancel.Visible = false;
            ToggleButtons(true);
            AppSettingsStore.Save(_settings);
        }
    }

    private bool IsEnglish => Strings.IsEnglish;

    private void ToggleButtons(bool enabled)
    {
        _btnSelectFolders.Enabled = enabled;
        _btnRefresh.Enabled = enabled;
        _btnRebuild.Enabled = enabled;
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
