using MaterialSkin;
using MaterialSkin.Controls;

namespace OutlookClassicSearch;

internal sealed class HelpForm : MaterialForm
{
    public HelpForm()
    {
        // MaterialSkin toevoegen aan dit formulier
        var materialSkinManager = MaterialSkinManager.Instance;
        materialSkinManager.AddFormToManage(this);

        Text = "Help";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(640, 430);

        AppVisualAssets.ApplyWindowIcon(this);

        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(14)
        };

        var logoBox = new PictureBox
        {
            Dock = DockStyle.Top,
            Height = 90,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        logoBox.Image = AppVisualAssets.TryLoadLogoImage();

        var titleLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 34,
            Text = "Help – " + Strings.InfoAppName,
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(0, 8, 0, 0)
        };

        var contentBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            WordWrap = true,
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 10),
            BackColor = SystemColors.Control,
            Text = Strings.HelpContent
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

        panel.Controls.Add(contentBox);
        panel.Controls.Add(titleLabel);
        panel.Controls.Add(logoBox);
        Controls.Add(panel);
        Controls.Add(bottomPanel);

        AcceptButton = btnClose;
        AppTheme.Apply(this);
    }
}
