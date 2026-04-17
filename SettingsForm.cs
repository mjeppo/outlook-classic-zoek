using MaterialSkin;
using MaterialSkin.Controls;

namespace OutlookClassicSearch;

internal partial class SettingsForm : MaterialForm
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

        // MaterialSkin toevoegen aan dit formulier
        var materialSkinManager = MaterialSkinManager.Instance;
        materialSkinManager.AddFormToManage(this);

        AppVisualAssets.ApplyWindowIcon(this);
        AppTheme.Apply(this);
        AppTheme.ApplyPrimaryStyle(btnOK);
        AppTheme.ApplySecondaryStyle(btnCancel);
        AppTheme.ApplySecondaryStyle(btnRefreshStores);
        AppTheme.ApplySecondaryStyle(btnChooseExcludedFolders);
        AppTheme.ApplySecondaryStyle(btnManageIndex);
        AppTheme.ApplySecondaryStyle(btnClearHistory);

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

        // Theme instelling
        cmbTheme.SelectedIndex = settings.Theme switch
        {
            "Dark" => 1,
            "Auto" => 2,
            _ => 0
        };

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

        // Pas het nieuwe thema toe
        AppTheme.ChangeTheme(_settings.Theme switch
        {
            "Dark" => AppTheme.ThemeMode.Dark,
            "Auto" => AppTheme.ThemeMode.Auto,
            _ => AppTheme.ThemeMode.Light
        });

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

        // Theme instelling
        _settings.Theme = cmbTheme.SelectedIndex switch
        {
            1 => "Dark",
            2 => "Auto",
            _ => "Light"
        };

        var checkedStores = GetCheckedStores();
        _settings.SelectedStoreIds = checkedStores.Select(s => s.StoreId).ToList();
        _settings.SelectedStoreDisplayNames = checkedStores.Select(s => s.DisplayName).ToList();
    }

    private void ApplyStrings()
    {
        Text = Strings.SettingsTitle;
        lblSearchBody.Text = Strings.SettingsChkSearchBody;
        lblSearchAttachments.Text = Strings.SettingsChkSearchAtt;
        lblUseIndex.Text = Strings.SettingsChkUseIndex;
        lblMaxResults.Text = Strings.SettingsLblMaxResults;
        lblExcludeAttachmentExt.Text = Strings.SettingsChkExcludeExt;
        lblStores.Text = Strings.SettingsLblStores;
        btnRefreshStores.Text = Strings.SettingsBtnRefreshStores;
        btnChooseExcludedFolders.Text = Strings.SettingsBtnExcludeFolders;
        btnManageIndex.Text = Strings.SettingsBtnManageIndex;
        lblSearchHistoryMax.Text = Strings.SettingsLblSearchHistoryMax;
        btnClearHistory.Text = Strings.SettingsBtnClearHistory;
        lblTheme.Text = Strings.SettingsLblTheme;
        lblLanguage.Text = Strings.SettingsLblLanguage;
        btnOK.Text = Strings.BtnSave;
        btnCancel.Text = Strings.BtnCancel;
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

    // Label click handlers om switches te togglen
    private void lblSearchBody_Click(object? sender, EventArgs e)
    {
        chkSearchBody.Checked = !chkSearchBody.Checked;
    }

    private void lblSearchAttachments_Click(object? sender, EventArgs e)
    {
        chkSearchAttachments.Checked = !chkSearchAttachments.Checked;
    }

    private void lblUseIndex_Click(object? sender, EventArgs e)
    {
        chkUseIndex.Checked = !chkUseIndex.Checked;
    }

    private void lblExcludeAttachmentExt_Click(object? sender, EventArgs e)
    {
        chkExcludeAttachmentExt.Checked = !chkExcludeAttachmentExt.Checked;
    }

    private void nudMaxResults_ValueChanged(object sender, EventArgs e)
    {

    }
}


