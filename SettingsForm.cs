namespace OutlookClassicSearch;

internal partial class SettingsForm : Form
{
    private readonly AppSettings _settings;
    private List<MailboxFolderRoot> _folderRoots;
    private List<string> _folderTreeStoreIds;

    public List<StoreListItem> AvailableStores { get; private set; }
    public List<MailboxFolderRoot> FolderRoots => _folderRoots;
    public List<string> FolderTreeStoreIds => _folderTreeStoreIds;

    public SettingsForm(
        AppSettings settings,
        List<StoreListItem> availableStores,
        List<MailboxFolderRoot> folderRoots,
        List<string> folderTreeStoreIds)
    {
        InitializeComponent();
        AppVisualAssets.ApplyWindowIcon(this);
        AppTheme.Apply(this);
        AppTheme.ApplyPrimaryStyle(btnOK);

        _settings = settings;
        AvailableStores = new List<StoreListItem>(availableStores);
        _folderRoots = folderRoots;
        _folderTreeStoreIds = folderTreeStoreIds;

        chkExcludeAttachmentExt.CheckedChanged += (_, _) => txtExcludedAttachmentExt.Enabled = chkExcludeAttachmentExt.Checked;

        chkSearchBody.Checked = settings.SearchBody;
        chkSearchAttachments.Checked = settings.SearchAttachments;
        nudMaxResults.Value = Math.Clamp(settings.MaxResults, 10, 5000);
        chkUseIndex.Checked = settings.UsePersistentIndexForSearch;
        chkExcludeAttachmentExt.Checked = settings.ExcludeAttachmentExtensions;
        txtExcludedAttachmentExt.Text = settings.ExcludedAttachmentExtensionsRaw;
        txtExcludedAttachmentExt.Enabled = settings.ExcludeAttachmentExtensions;
        lblExcludedFolderSummary.Text = string.Format(Strings.SettingsLblExcludedFoldersFmt, settings.ExcludedFolderEntryIds.Count);
        cmbLanguage.SelectedIndex = settings.Language == "en" ? 1 : 0;
    nudSearchHistoryMax.Value = Math.Clamp(settings.SearchHistoryMaxCount, 1, 50);

        ApplyStrings();
        PopulateStores();
    }

    private void PopulateStores()
    {
        clbStores.Items.Clear();
        var selectedIds = new HashSet<string>(_settings.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
        var selectedNames = new HashSet<string>(_settings.SelectedStoreDisplayNames, StringComparer.OrdinalIgnoreCase);

        foreach (var store in AvailableStores)
        {
            bool hasSaved = selectedIds.Count > 0 || selectedNames.Count > 0;
            bool isChecked = !hasSaved || selectedIds.Contains(store.StoreId) || selectedNames.Contains(store.DisplayName);
            clbStores.Items.Add(store, isChecked);
        }
    }

    private async void btnRefreshStores_Click(object sender, EventArgs e)
    {
        btnRefreshStores.Enabled = false;
        lblStoresStatus.Text = Strings.SettingsMsgLoadingMailboxes;
        try
        {
            var stores = await OutlookSearcher.GetAvailableStoresAsync(CancellationToken.None);
            AvailableStores = stores.Select(s => new StoreListItem(s.DisplayName, s.StoreId)).ToList();
            PopulateStores();
            lblStoresStatus.Text = string.Format(Strings.StatusMailboxesFoundFmt, AvailableStores.Count);
        }
        catch (Exception ex)
        {
            lblStoresStatus.Text = Strings.SettingsMsgMailboxesFailed;
            ErrorDetailsDialog.Show(this, Strings.ErrMailboxSettingsTitle, Strings.ErrMailboxSettingsSummary, ex);
        }
        finally
        {
            btnRefreshStores.Enabled = true;
        }
    }

    private async void btnChooseExcludedFolders_Click(object sender, EventArgs e)
    {
        var selectedStores = GetCheckedStores();
        if (selectedStores.Count == 0)
        {
            MessageBox.Show(this, Strings.SettingsMsgNoMailbox, Strings.SettingsMsgNoMailboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                lblStoresStatus.Text = Strings.SettingsMsgLoadingFolders;
                var roots = await OutlookSearcher.GetFolderTreeAsync(selectedStoreIds, cancellationToken);
                _folderRoots = roots.ToList();
                _folderTreeStoreIds = selectedStoreIds;
                lblStoresStatus.Text = string.Empty;
                return roots;
            });

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _settings.ExcludedFolderEntryIds = dialog.GetSelectedEntryIds().ToList();
            lblExcludedFolderSummary.Text = string.Format(Strings.SettingsLblExcludedFoldersFmt, _settings.ExcludedFolderEntryIds.Count);
        }
    }

    private void btnManageIndex_Click(object sender, EventArgs e)
    {
        var selectedStores = GetCheckedStores();
        using var form = new IndexManagerForm(selectedStores, _settings, _folderRoots);
        form.ShowDialog(this);
    }

    private void btnOK_Click(object sender, EventArgs e)
    {
        SaveToSettings();
        DialogResult = DialogResult.OK;
        Close();
    }

    private void btnCancel_Click(object sender, EventArgs e)
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private void SaveToSettings()
    {
        _settings.SearchBody = chkSearchBody.Checked;
        _settings.SearchAttachments = chkSearchAttachments.Checked;
        _settings.MaxResults = (int)nudMaxResults.Value;
        _settings.UsePersistentIndexForSearch = chkUseIndex.Checked;
        _settings.ExcludeAttachmentExtensions = chkExcludeAttachmentExt.Checked;
        _settings.ExcludedAttachmentExtensionsRaw = txtExcludedAttachmentExt.Text.Trim();
        _settings.Language = cmbLanguage.SelectedIndex == 1 ? "en" : "nl";
        Strings.IsEnglish = _settings.Language == "en";
    _settings.SearchHistoryMaxCount = (int)nudSearchHistoryMax.Value;

        var checkedStores = GetCheckedStores();
        _settings.SelectedStoreIds = checkedStores.Select(s => s.StoreId).ToList();
        _settings.SelectedStoreDisplayNames = checkedStores.Select(s => s.DisplayName).ToList();
    }

    private void ApplyStrings()
    {
        Text = Strings.SettingsTitle;
        chkSearchBody.Text             = Strings.SettingsChkSearchBody;
        chkSearchAttachments.Text      = Strings.SettingsChkSearchAtt;
        chkUseIndex.Text               = Strings.SettingsChkUseIndex;
        lblMaxResults.Text             = Strings.SettingsLblMaxResults;
        chkExcludeAttachmentExt.Text   = Strings.SettingsChkExcludeExt;
        lblStores.Text                 = Strings.SettingsLblStores;
        btnRefreshStores.Text          = Strings.SettingsBtnRefreshStores;
        btnChooseExcludedFolders.Text   = Strings.SettingsBtnExcludeFolders;
        btnManageIndex.Text            = Strings.SettingsBtnManageIndex;
        lblLanguage.Text               = Strings.SettingsLblLanguage;
        lblSearchHistoryMax.Text       = Strings.SettingsLblSearchHistoryMax;
        btnClearHistory.Text           = Strings.SettingsBtnClearHistory;
        btnOK.Text                     = Strings.BtnSave;
        btnCancel.Text                 = Strings.BtnCancel;
    }


    private void btnClearHistory_Click(object sender, EventArgs e)
    {
        _settings.SearchHistory.Clear();
        AppSettingsStore.Save(_settings);
    }

    private List<StoreListItem> GetCheckedStores()
    {
        return clbStores.CheckedItems.OfType<StoreListItem>().ToList();
    }
}


