namespace OutlookClassicSearch;

partial class SettingsForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
            components.Dispose();
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        chkSearchBody = new MaterialSkin.Controls.MaterialSwitch();
        lblSearchBody = new Label();
        chkSearchAttachments = new MaterialSkin.Controls.MaterialSwitch();
        lblSearchAttachments = new Label();
        chkUseIndex = new MaterialSkin.Controls.MaterialSwitch();
        lblUseIndex = new Label();
        lblMaxResults = new Label();
        nudMaxResults = new NumericUpDown();
        chkExcludeAttachmentExt = new MaterialSkin.Controls.MaterialSwitch();
        lblExcludeAttachmentExt = new Label();
        txtExcludedAttachmentExt = new TextBox();
        lblStores = new Label();
        btnRefreshStores = new Button();
        clbStores = new CheckedListBox();
        lblStoresStatus = new Label();
        btnChooseExcludedFolders = new Button();
        lblExcludedFolderSummary = new Label();
        btnManageIndex = new Button();
        lblSearchHistoryMax = new Label();
        nudSearchHistoryMax = new NumericUpDown();
        btnClearHistory = new Button();
        lblWarningThreshold = new Label();
        nudWarningThreshold = new NumericUpDown();
        lblTheme = new Label();
        cmbTheme = new ComboBox();
        lblLanguage = new Label();
        cmbLanguage = new ComboBox();
        btnOK = new Button();
        btnApply = new Button();
        btnCancel = new Button();
        lblExcludeAttachmentExt1 = new Label();
        lblBijschriftBerichten = new Label();
        lblMappenUitsluiten = new Label();
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudSearchHistoryMax).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudWarningThreshold).BeginInit();
        SuspendLayout();
        // 
        // chkSearchBody
        // 
        chkSearchBody.AutoSize = true;
        chkSearchBody.Depth = 0;
        chkSearchBody.Location = new Point(12, 76);
        chkSearchBody.Margin = new Padding(0);
        chkSearchBody.MouseLocation = new Point(-1, -1);
        chkSearchBody.MouseState = MaterialSkin.MouseState.HOVER;
        chkSearchBody.Name = "chkSearchBody";
        chkSearchBody.Ripple = true;
        chkSearchBody.Size = new Size(58, 37);
        chkSearchBody.TabIndex = 0;
        chkSearchBody.UseVisualStyleBackColor = true;
        // 
        // lblSearchBody
        // 
        lblSearchBody.AutoSize = true;
        lblSearchBody.Cursor = Cursors.Hand;
        lblSearchBody.Location = new Point(74, 85);
        lblSearchBody.Name = "lblSearchBody";
        lblSearchBody.Size = new Size(99, 15);
        lblSearchBody.TabIndex = 100;
        lblSearchBody.Text = "Zoek ook in body";
        lblSearchBody.Click += lblSearchBody_Click;
        // 
        // chkSearchAttachments
        // 
        chkSearchAttachments.AutoSize = true;
        chkSearchAttachments.Depth = 0;
        chkSearchAttachments.Location = new Point(230, 76);
        chkSearchAttachments.Margin = new Padding(0);
        chkSearchAttachments.MouseLocation = new Point(-1, -1);
        chkSearchAttachments.MouseState = MaterialSkin.MouseState.HOVER;
        chkSearchAttachments.Name = "chkSearchAttachments";
        chkSearchAttachments.Ripple = true;
        chkSearchAttachments.Size = new Size(58, 37);
        chkSearchAttachments.TabIndex = 1;
        chkSearchAttachments.UseVisualStyleBackColor = true;
        // 
        // lblSearchAttachments
        // 
        lblSearchAttachments.AutoSize = true;
        lblSearchAttachments.Cursor = Cursors.Hand;
        lblSearchAttachments.Location = new Point(292, 85);
        lblSearchAttachments.Name = "lblSearchAttachments";
        lblSearchAttachments.Size = new Size(114, 15);
        lblSearchAttachments.TabIndex = 101;
        lblSearchAttachments.Text = "Zoek ook in bijlagen";
        lblSearchAttachments.Click += lblSearchAttachments_Click;
        // 
        // chkUseIndex
        // 
        chkUseIndex.AutoSize = true;
        chkUseIndex.Depth = 0;
        chkUseIndex.Location = new Point(480, 76);
        chkUseIndex.Margin = new Padding(0);
        chkUseIndex.MouseLocation = new Point(-1, -1);
        chkUseIndex.MouseState = MaterialSkin.MouseState.HOVER;
        chkUseIndex.Name = "chkUseIndex";
        chkUseIndex.Ripple = true;
        chkUseIndex.Size = new Size(58, 37);
        chkUseIndex.TabIndex = 2;
        chkUseIndex.UseVisualStyleBackColor = true;
        // 
        // lblUseIndex
        // 
        lblUseIndex.AutoSize = true;
        lblUseIndex.Cursor = Cursors.Hand;
        lblUseIndex.Location = new Point(542, 85);
        lblUseIndex.Name = "lblUseIndex";
        lblUseIndex.Size = new Size(146, 15);
        lblUseIndex.TabIndex = 102;
        lblUseIndex.Text = "Gebruik permanente index";
        lblUseIndex.Click += lblUseIndex_Click;
        // 
        // lblMaxResults
        // 
        lblMaxResults.AutoSize = true;
        lblMaxResults.Location = new Point(12, 489);
        lblMaxResults.Name = "lblMaxResults";
        lblMaxResults.Size = new Size(111, 15);
        lblMaxResults.TabIndex = 36;
        lblMaxResults.Text = "Max zoekresultaten:";
        // 
        // nudMaxResults
        // 
        nudMaxResults.Location = new Point(214, 487);
        nudMaxResults.Maximum = new decimal(new int[] { 5000, 0, 0, 0 });
        nudMaxResults.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
        nudMaxResults.Name = "nudMaxResults";
        nudMaxResults.Size = new Size(86, 23);
        nudMaxResults.TabIndex = 35;
        nudMaxResults.Value = new decimal(new int[] { 5000, 0, 0, 0 });
        nudMaxResults.ValueChanged += nudMaxResults_ValueChanged;
        // 
        // chkExcludeAttachmentExt
        // 
        chkExcludeAttachmentExt.AutoSize = true;
        chkExcludeAttachmentExt.Depth = 0;
        chkExcludeAttachmentExt.Location = new Point(12, 115);
        chkExcludeAttachmentExt.Margin = new Padding(0);
        chkExcludeAttachmentExt.MouseLocation = new Point(-1, -1);
        chkExcludeAttachmentExt.MouseState = MaterialSkin.MouseState.HOVER;
        chkExcludeAttachmentExt.Name = "chkExcludeAttachmentExt";
        chkExcludeAttachmentExt.Ripple = true;
        chkExcludeAttachmentExt.Size = new Size(58, 37);
        chkExcludeAttachmentExt.TabIndex = 3;
        chkExcludeAttachmentExt.UseVisualStyleBackColor = true;
        // 
        // lblExcludeAttachmentExt
        // 
        lblExcludeAttachmentExt.AutoSize = true;
        lblExcludeAttachmentExt.Cursor = Cursors.Hand;
        lblExcludeAttachmentExt.Location = new Point(74, 147);
        lblExcludeAttachmentExt.Name = "lblExcludeAttachmentExt";
        lblExcludeAttachmentExt.Size = new Size(203, 15);
        lblExcludeAttachmentExt.TabIndex = 103;
        lblExcludeAttachmentExt.Text = "Bijlage-extensies uitsluiten (;):";
        lblExcludeAttachmentExt.Click += lblExcludeAttachmentExt_Click;
        // 
        // txtExcludedAttachmentExt
        // 
        txtExcludedAttachmentExt.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtExcludedAttachmentExt.Location = new Point(12, 147);
        txtExcludedAttachmentExt.Name = "txtExcludedAttachmentExt";
        txtExcludedAttachmentExt.PlaceholderText = ".zip;.png;.jpg;.jpeg;.gif";
        txtExcludedAttachmentExt.Size = new Size(738, 23);
        txtExcludedAttachmentExt.TabIndex = 34;
        // 
        // lblStores
        // 
        lblStores.AutoSize = true;
        lblStores.Location = new Point(12, 198);
        lblStores.Name = "lblStores";
        lblStores.Size = new Size(136, 15);
        lblStores.TabIndex = 33;
        lblStores.Text = "Mailboxen (incl. shared):";
        // 
        // btnRefreshStores
        // 
        btnRefreshStores.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnRefreshStores.Location = new Point(579, 194);
        btnRefreshStores.Name = "btnRefreshStores";
        btnRefreshStores.Size = new Size(171, 23);
        btnRefreshStores.TabIndex = 32;
        btnRefreshStores.Text = "Mailboxen vernieuwen";
        btnRefreshStores.UseVisualStyleBackColor = true;
        btnRefreshStores.Click += btnRefreshStores_Click;
        // 
        // clbStores
        // 
        clbStores.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        clbStores.CheckOnClick = true;
        clbStores.FormattingEnabled = true;
        clbStores.Location = new Point(12, 220);
        clbStores.Name = "clbStores";
        clbStores.Size = new Size(738, 184);
        clbStores.TabIndex = 10;
        // 
        // lblStoresStatus
        // 
        lblStoresStatus.AutoSize = true;
        lblStoresStatus.Location = new Point(12, 386);
        lblStoresStatus.Name = "lblStoresStatus";
        lblStoresStatus.Size = new Size(0, 15);
        lblStoresStatus.TabIndex = 31;
        // 
        // btnChooseExcludedFolders
        // 
        btnChooseExcludedFolders.Location = new Point(214, 422);
        btnChooseExcludedFolders.Name = "btnChooseExcludedFolders";
        btnChooseExcludedFolders.Size = new Size(111, 27);
        btnChooseExcludedFolders.TabIndex = 30;
        btnChooseExcludedFolders.Text = "Kies map(pen)";
        btnChooseExcludedFolders.UseVisualStyleBackColor = true;
        btnChooseExcludedFolders.Click += btnChooseExcludedFolders_Click;
        // 
        // lblExcludedFolderSummary
        // 
        lblExcludedFolderSummary.AutoSize = true;
        lblExcludedFolderSummary.Location = new Point(340, 428);
        lblExcludedFolderSummary.Name = "lblExcludedFolderSummary";
        lblExcludedFolderSummary.Size = new Size(126, 15);
        lblExcludedFolderSummary.TabIndex = 29;
        lblExcludedFolderSummary.Text = "Uitgesloten mappen: 0";
        // 
        // btnManageIndex
        // 
        btnManageIndex.Location = new Point(578, 416);
        btnManageIndex.Name = "btnManageIndex";
        btnManageIndex.Size = new Size(140, 27);
        btnManageIndex.TabIndex = 28;
        btnManageIndex.Text = "Index beheren...";
        btnManageIndex.UseVisualStyleBackColor = true;
        btnManageIndex.Click += btnManageIndex_Click;
        // 
        // lblSearchHistoryMax
        // 
        lblSearchHistoryMax.AutoSize = true;
        lblSearchHistoryMax.Location = new Point(12, 457);
        lblSearchHistoryMax.Name = "lblSearchHistoryMax";
        lblSearchHistoryMax.Size = new Size(127, 15);
        lblSearchHistoryMax.TabIndex = 27;
        lblSearchHistoryMax.Text = "Max zoekgeschiedenis:";
        // 
        // nudSearchHistoryMax
        // 
        nudSearchHistoryMax.Location = new Point(214, 455);
        nudSearchHistoryMax.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
        nudSearchHistoryMax.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
        nudSearchHistoryMax.Name = "nudSearchHistoryMax";
        nudSearchHistoryMax.Size = new Size(60, 23);
        nudSearchHistoryMax.TabIndex = 21;
        nudSearchHistoryMax.Value = new decimal(new int[] { 10, 0, 0, 0 });
        // 
        // btnClearHistory
        // 
        btnClearHistory.Location = new Point(291, 453);
        btnClearHistory.Name = "btnClearHistory";
        btnClearHistory.Size = new Size(160, 27);
        btnClearHistory.TabIndex = 26;
        btnClearHistory.Text = "Geschiedenis wissen";
        btnClearHistory.UseVisualStyleBackColor = true;
        btnClearHistory.Click += btnClearHistory_Click;
        // 
        // lblWarningThreshold
        // 
        lblWarningThreshold.AutoSize = true;
        lblWarningThreshold.Location = new Point(12, 522);
        lblWarningThreshold.Name = "lblWarningThreshold";
        lblWarningThreshold.Size = new Size(196, 15);
        lblWarningThreshold.TabIndex = 29;
        lblWarningThreshold.Text = "Waarschuwing te veel resultaten bij:";
        // 
        // nudWarningThreshold
        // 
        nudWarningThreshold.Location = new Point(214, 520);
        nudWarningThreshold.Maximum = new decimal(new int[] { 10000, 0, 0, 0 });
        nudWarningThreshold.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
        nudWarningThreshold.Name = "nudWarningThreshold";
        nudWarningThreshold.Size = new Size(62, 23);
        nudWarningThreshold.TabIndex = 28;
        nudWarningThreshold.Value = new decimal(new int[] { 100, 0, 0, 0 });
        // 
        // lblTheme
        // 
        lblTheme.AutoSize = true;
        lblTheme.Location = new Point(12, 557);
        lblTheme.Name = "lblTheme";
        lblTheme.Size = new Size(47, 15);
        lblTheme.TabIndex = 25;
        lblTheme.Text = "Thema:";
        // 
        // cmbTheme
        // 
        cmbTheme.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbTheme.Items.AddRange(new object[] { "Licht", "Donker", "Auto (Systeem)" });
        cmbTheme.Location = new Point(214, 554);
        cmbTheme.Name = "cmbTheme";
        cmbTheme.Size = new Size(140, 23);
        cmbTheme.TabIndex = 22;
        // 
        // lblLanguage
        // 
        lblLanguage.AutoSize = true;
        lblLanguage.Location = new Point(12, 591);
        lblLanguage.Name = "lblLanguage";
        lblLanguage.Size = new Size(81, 15);
        lblLanguage.TabIndex = 24;
        lblLanguage.Text = "Weergavetaal:";
        // 
        // cmbLanguage
        // 
        cmbLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbLanguage.Items.AddRange(new object[] { "Nederlands", "English" });
        cmbLanguage.Location = new Point(214, 588);
        cmbLanguage.Name = "cmbLanguage";
        cmbLanguage.Size = new Size(140, 23);
        cmbLanguage.TabIndex = 23;
        // 
        // btnOK
        // 
        btnOK.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnOK.Location = new Point(498, 679);
        btnOK.Name = "btnOK";
        btnOK.Size = new Size(80, 27);
        btnOK.TabIndex = 1;
        btnOK.Text = "Opslaan";
        btnOK.UseVisualStyleBackColor = true;
        btnOK.Click += btnOK_Click;
        // 
        // btnApply
        // 
        btnApply.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnApply.Location = new Point(588, 679);
        btnApply.Name = "btnApply";
        btnApply.Size = new Size(80, 27);
        btnApply.TabIndex = 2;
        btnApply.Text = "Toepassen";
        btnApply.UseVisualStyleBackColor = true;
        btnApply.Click += btnApply_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.Location = new Point(678, 679);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(80, 27);
        btnCancel.TabIndex = 0;
        btnCancel.Text = "Annuleren";
        btnCancel.UseVisualStyleBackColor = true;
        btnCancel.Click += btnCancel_Click;
        // 
        // lblExcludeAttachmentExt1
        // 
        lblExcludeAttachmentExt1.AutoSize = true;
        lblExcludeAttachmentExt1.Cursor = Cursors.Hand;
        lblExcludeAttachmentExt1.Location = new Point(73, 125);
        lblExcludeAttachmentExt1.Name = "lblExcludeAttachmentExt1";
        lblExcludeAttachmentExt1.Size = new Size(284, 15);
        lblExcludeAttachmentExt1.TabIndex = 103;
        lblExcludeAttachmentExt1.Text = "Bijlage-extensies uitsluiten (woorden scheiden met ;)";
        // 
        // lblBijschriftBerichten
        // 
        lblBijschriftBerichten.AutoSize = true;
        lblBijschriftBerichten.Location = new Point(291, 522);
        lblBijschriftBerichten.Name = "lblBijschriftBerichten";
        lblBijschriftBerichten.Size = new Size(57, 15);
        lblBijschriftBerichten.TabIndex = 104;
        lblBijschriftBerichten.Text = "berichten";
        // 
        // lblMappenUitsluiten
        // 
        lblMappenUitsluiten.AutoSize = true;
        lblMappenUitsluiten.Location = new Point(12, 428);
        lblMappenUitsluiten.Name = "lblMappenUitsluiten";
        lblMappenUitsluiten.Size = new Size(162, 15);
        lblMappenUitsluiten.TabIndex = 105;
        lblMappenUitsluiten.Text = "Mappen uitsluiten bij zoeken:";
        // 
        // SettingsForm
        // 
        AcceptButton = btnOK;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(770, 719);
        Controls.Add(lblMappenUitsluiten);
        Controls.Add(lblBijschriftBerichten);
        Controls.Add(lblExcludeAttachmentExt1);
        Controls.Add(btnCancel);
        Controls.Add(btnApply);
        Controls.Add(btnOK);
        Controls.Add(cmbLanguage);
        Controls.Add(lblLanguage);
        Controls.Add(cmbTheme);
        Controls.Add(lblTheme);
        Controls.Add(nudWarningThreshold);
        Controls.Add(lblWarningThreshold);
        Controls.Add(btnClearHistory);
        Controls.Add(nudSearchHistoryMax);
        Controls.Add(lblSearchHistoryMax);
        Controls.Add(btnManageIndex);
        Controls.Add(lblExcludedFolderSummary);
        Controls.Add(btnChooseExcludedFolders);
        Controls.Add(lblStoresStatus);
        Controls.Add(clbStores);
        Controls.Add(btnRefreshStores);
        Controls.Add(lblStores);
        Controls.Add(txtExcludedAttachmentExt);
        Controls.Add(chkExcludeAttachmentExt);
        Controls.Add(nudMaxResults);
        Controls.Add(lblMaxResults);
        Controls.Add(lblUseIndex);
        Controls.Add(chkUseIndex);
        Controls.Add(lblSearchAttachments);
        Controls.Add(chkSearchAttachments);
        Controls.Add(lblSearchBody);
        Controls.Add(chkSearchBody);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "SettingsForm";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Instellingen";
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudSearchHistoryMax).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudWarningThreshold).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private MaterialSkin.Controls.MaterialSwitch chkSearchBody;
    private Label lblSearchBody;
    private MaterialSkin.Controls.MaterialSwitch chkSearchAttachments;
    private Label lblSearchAttachments;
    private MaterialSkin.Controls.MaterialSwitch chkUseIndex;
    private Label lblUseIndex;
    private Label lblMaxResults;
    private NumericUpDown nudMaxResults;
    private MaterialSkin.Controls.MaterialSwitch chkExcludeAttachmentExt;
    private Label lblExcludeAttachmentExt;
    private TextBox txtExcludedAttachmentExt;
    private Label lblStores;
    private Button btnRefreshStores;
    private CheckedListBox clbStores;
    private Label lblStoresStatus;
    private Button btnChooseExcludedFolders;
    private Label lblExcludedFolderSummary;
    private Button btnManageIndex;
    private Label lblSearchHistoryMax;
    private NumericUpDown nudSearchHistoryMax;
    private Button btnClearHistory;
    private Label lblWarningThreshold;
    private NumericUpDown nudWarningThreshold;
    private Label lblTheme;
    private ComboBox cmbTheme;
    private Label lblLanguage;
    private ComboBox cmbLanguage;
    private Button btnOK;
    private Button btnApply;
    private Button btnCancel;
    private Label lblExcludeAttachmentExt1;
    private Label lblBijschriftBerichten;
    private Label lblMappenUitsluiten;
}
