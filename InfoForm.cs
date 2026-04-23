using MaterialSkin;
using MaterialSkin.Controls;

namespace OutlookClassicSearch;

internal sealed class InfoForm : MaterialForm
{
    public InfoForm()
    {
        // MaterialSkin toevoegen aan dit formulier
        var materialSkinManager = MaterialSkinManager.Instance;
        materialSkinManager.AddFormToManage(this);

        string version = System.Reflection.Assembly.GetExecutingAssembly()
            .GetName().Version?.ToString(3) ?? "onbekend";

        Text = "Info";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(640, 450);

        AppVisualAssets.ApplyWindowIcon(this);

        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(14)
        };

        var logoBox = new PictureBox
        {
            Dock = DockStyle.Top,
            Height = 135,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };

        logoBox.Image = AppVisualAssets.TryLoadLogoImage();

        var titleLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 34,
            Text = Strings.InfoAppName,
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(0, 10, 0, 0)
        };

        var infoLabel = new Label
        {
            Dock = DockStyle.Fill,
            Text =
                string.Format(Strings.InfoVersionFmt, version) + "\n\n" +
                Strings.InfoDesc + "\n\n" +
                Strings.InfoBuiltWith,
            Font = new Font("Segoe UI", 10),
            TextAlign = ContentAlignment.TopLeft
        };

        var btnClose = new Button
        {
            Text = Strings.BtnClose,
            Width = 100,
            Height = 30,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
            DialogResult = DialogResult.OK
        };
        AppTheme.ApplyPrimaryStyle(btnClose);

        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 40
        };
        btnClose.Location = new Point(bottomPanel.Width - btnClose.Width, 5);
        bottomPanel.Resize += (_, _) => btnClose.Left = bottomPanel.ClientSize.Width - btnClose.Width;
        bottomPanel.Controls.Add(btnClose);

        panel.Controls.Add(infoLabel);
        panel.Controls.Add(titleLabel);
        panel.Controls.Add(logoBox);
        Controls.Add(panel);
        Controls.Add(bottomPanel);

        AcceptButton = btnClose;
        AppTheme.Apply(this);
    }
}
