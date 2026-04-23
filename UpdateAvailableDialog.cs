using MaterialSkin.Controls;

namespace OutlookClassicSearch;

/// <summary>
/// Dialog die update informatie toont
/// </summary>
internal sealed class UpdateAvailableDialog : MaterialForm
{
    private readonly UpdateInfo _updateInfo;
    private readonly Label _lblTitle = new();
    private readonly Label _lblCurrentVersion = new();
    private readonly Label _lblNewVersion = new();
    private readonly Label _lblReleaseDate = new();
    private readonly TextBox _txtReleaseNotes = new();
    private readonly Button _btnDownload = new();
    private readonly Button _btnLater = new();

    public UpdateAvailableDialog(UpdateInfo updateInfo)
    {
        _updateInfo = updateInfo;

        // MaterialSkin toevoegen
        var materialSkinManager = MaterialSkin.MaterialSkinManager.Instance;
        materialSkinManager.AddFormToManage(this);

        Text = Strings.UpdateTitle;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(500, 400);
        AppVisualAssets.ApplyWindowIcon(this);

        // Title label
        _lblTitle.Text = Strings.UpdateAvailableTitle;
        _lblTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
        _lblTitle.SetBounds(20, 80, 460, 30);

        // Current version
        _lblCurrentVersion.Text = string.Format(Strings.UpdateCurrentVersionFmt, _updateInfo.CurrentVersion);
        _lblCurrentVersion.SetBounds(20, 120, 460, 20);

        // New version
        _lblNewVersion.Text = string.Format(Strings.UpdateNewVersionFmt, _updateInfo.LatestVersion);
        _lblNewVersion.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        _lblNewVersion.SetBounds(20, 145, 460, 20);

        // Release date
        _lblReleaseDate.Text = string.Format(Strings.UpdateReleaseDateFmt, _updateInfo.PublishedAt.ToLocalTime().ToString("d MMMM yyyy"));
        _lblReleaseDate.SetBounds(20, 170, 460, 20);

        // Release notes
        var lblNotes = new Label
        {
            Text = Strings.UpdateReleaseNotes,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold)
        };
        lblNotes.SetBounds(20, 200, 460, 20);

        _txtReleaseNotes.Multiline = true;
        _txtReleaseNotes.ReadOnly = true;
        _txtReleaseNotes.ScrollBars = ScrollBars.Vertical;
        _txtReleaseNotes.Text = _updateInfo.ReleaseNotes;
        _txtReleaseNotes.SetBounds(20, 225, 460, 110);
        _txtReleaseNotes.BackColor = SystemColors.Window;

        // Buttons
        _btnDownload.Text = Strings.UpdateBtnDownload;
        _btnDownload.SetBounds(270, 350, 100, 32);
        _btnDownload.Click += (_, _) =>
        {
            DialogResult = DialogResult.OK;
            Close();
        };

        _btnLater.Text = Strings.UpdateBtnLater;
        _btnLater.SetBounds(380, 350, 100, 32);
        _btnLater.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };

        Controls.Add(_lblTitle);
        Controls.Add(_lblCurrentVersion);
        Controls.Add(_lblNewVersion);
        Controls.Add(_lblReleaseDate);
        Controls.Add(lblNotes);
        Controls.Add(_txtReleaseNotes);
        Controls.Add(_btnDownload);
        Controls.Add(_btnLater);

        AppTheme.Apply(this);
        AppTheme.ApplyPrimaryStyle(_btnDownload);
    }
}
