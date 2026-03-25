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
        chkSearchBody = new CheckBox();
        chkSearchAttachments = new CheckBox();
        chkUseIndex = new CheckBox();
        lblMaxResults = new Label();
        nudMaxResults = new NumericUpDown();
        chkExcludeAttachmentExt = new CheckBox();
        txtExcludedAttachmentExt = new TextBox();
        lblStores = new Label();
        btnRefreshStores = new Button();
        clbStores = new CheckedListBox();
        lblStoresStatus = new Label();
        btnChooseExcludedFolders = new Button();
        lblExcludedFolderSummary = new Label();
        btnManageIndex = new Button();
        lblLanguage = new Label();
        cmbLanguage = new ComboBox();
        btnOK = new Button();
        btnCancel = new Button();
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).BeginInit();
            lblSearchHistoryMax = new Label();
            nudSearchHistoryMax = new NumericUpDown();
            btnClearHistory = new Button();
            ((System.ComponentModel.ISupportInitialize)nudSearchHistoryMax).BeginInit();
        SuspendLayout();
        // 
        // chkSearchBody
        // 
        chkSearchBody.AutoSize = true;
        chkSearchBody.Location = new Point(12, 16);
        chkSearchBody.Name = "chkSearchBody";
        chkSearchBody.Text = "Zoek ook in body";
        chkSearchBody.UseVisualStyleBackColor = true;
        // 
        // chkSearchAttachments
        // 
        chkSearchAttachments.AutoSize = true;
        chkSearchAttachments.Location = new Point(170, 16);
        chkSearchAttachments.Name = "chkSearchAttachments";
        chkSearchAttachments.Text = "Zoek ook in bijlagen";
        chkSearchAttachments.UseVisualStyleBackColor = true;
        // 
        // chkUseIndex
        // 
        chkUseIndex.AutoSize = true;
        chkUseIndex.Location = new Point(340, 16);
        chkUseIndex.Name = "chkUseIndex";
        chkUseIndex.Text = "Gebruik permanente index";
        chkUseIndex.UseVisualStyleBackColor = true;
        // 
        // lblMaxResults
        // 
        lblMaxResults.AutoSize = true;
        lblMaxResults.Location = new Point(12, 48);
        lblMaxResults.Name = "lblMaxResults";
        lblMaxResults.Text = "Max regels:";
        // 
        // nudMaxResults
        // 
        nudMaxResults.Location = new Point(88, 45);
        nudMaxResults.Maximum = new decimal(new int[] { 5000, 0, 0, 0 });
        nudMaxResults.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
        nudMaxResults.Name = "nudMaxResults";
        nudMaxResults.Size = new Size(86, 23);
        nudMaxResults.Value = new decimal(new int[] { 500, 0, 0, 0 });
        // 
        // chkExcludeAttachmentExt
        // 
        chkExcludeAttachmentExt.AutoSize = true;
        chkExcludeAttachmentExt.Location = new Point(12, 78);
        chkExcludeAttachmentExt.Name = "chkExcludeAttachmentExt";
        chkExcludeAttachmentExt.Text = "Bijlage-extensies uitsluiten (;):";
        chkExcludeAttachmentExt.UseVisualStyleBackColor = true;
        // 
        // txtExcludedAttachmentExt
        // 
        txtExcludedAttachmentExt.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtExcludedAttachmentExt.Location = new Point(12, 100);
        txtExcludedAttachmentExt.Name = "txtExcludedAttachmentExt";
        txtExcludedAttachmentExt.PlaceholderText = ".zip;.png;.jpg;.jpeg;.gif";
        txtExcludedAttachmentExt.Size = new Size(658, 23);
        // 
        // lblStores
        // 
        lblStores.AutoSize = true;
        lblStores.Location = new Point(12, 138);
        lblStores.Name = "lblStores";
        lblStores.Text = "Mailboxen (incl. shared):";
        // 
        // btnRefreshStores
        // 
        btnRefreshStores.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnRefreshStores.Location = new Point(499, 134);
        btnRefreshStores.Name = "btnRefreshStores";
        btnRefreshStores.Size = new Size(171, 23);
        btnRefreshStores.Text = "Mailboxen vernieuwen";
        btnRefreshStores.UseVisualStyleBackColor = true;
        btnRefreshStores.Click += btnRefreshStores_Click;
        // 
        // clbStores
        // 
        clbStores.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        clbStores.CheckOnClick = true;
        clbStores.FormattingEnabled = true;
        clbStores.Location = new Point(12, 160);
        clbStores.Name = "clbStores";
        clbStores.Size = new Size(658, 160);
        clbStores.TabIndex = 10;
        // 
        // lblStoresStatus
        // 
        lblStoresStatus.AutoSize = true;
        lblStoresStatus.Location = new Point(12, 326);
        lblStoresStatus.Name = "lblStoresStatus";
        lblStoresStatus.Text = string.Empty;
        // 
        // btnChooseExcludedFolders
        // 
        btnChooseExcludedFolders.Location = new Point(12, 356);
        btnChooseExcludedFolders.Name = "btnChooseExcludedFolders";
        btnChooseExcludedFolders.Size = new Size(210, 27);
        btnChooseExcludedFolders.Text = "Mappen uitsluiten van zoeken...";
        btnChooseExcludedFolders.UseVisualStyleBackColor = true;
        btnChooseExcludedFolders.Click += btnChooseExcludedFolders_Click;
        // 
        // lblExcludedFolderSummary
        // 
        lblExcludedFolderSummary.AutoSize = true;
        lblExcludedFolderSummary.Location = new Point(232, 362);
        lblExcludedFolderSummary.Name = "lblExcludedFolderSummary";
        lblExcludedFolderSummary.Text = "Uitgesloten mappen: 0";
        // 
        // btnManageIndex
        // 
        btnManageIndex.Location = new Point(12, 392);
        btnManageIndex.Name = "btnManageIndex";
        btnManageIndex.Size = new Size(140, 27);
        btnManageIndex.Text = "Index beheren...";
        btnManageIndex.UseVisualStyleBackColor = true;
        btnManageIndex.Click += btnManageIndex_Click;
        // 
        // lblLanguage
            // 
            // lblSearchHistoryMax
            // 
            lblSearchHistoryMax.AutoSize = true;
            lblSearchHistoryMax.Location = new Point(12, 432);
            lblSearchHistoryMax.Name = "lblSearchHistoryMax";
            lblSearchHistoryMax.Text = "Max zoekgeschiedenis:";
            // 
            // nudSearchHistoryMax
            // 
            nudSearchHistoryMax.Location = new Point(175, 428);
            nudSearchHistoryMax.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            nudSearchHistoryMax.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            nudSearchHistoryMax.Name = "nudSearchHistoryMax";
            nudSearchHistoryMax.Size = new Size(60, 23);
            nudSearchHistoryMax.Value = new decimal(new int[] { 10, 0, 0, 0 });
            nudSearchHistoryMax.TabIndex = 21;
            // 
            // btnClearHistory
            // 
            btnClearHistory.Location = new Point(252, 426);
            btnClearHistory.Name = "btnClearHistory";
            btnClearHistory.Size = new Size(160, 27);
            btnClearHistory.Text = "Geschiedenis wissen";
            btnClearHistory.UseVisualStyleBackColor = true;
            btnClearHistory.Click += btnClearHistory_Click;
        // 
        lblLanguage.AutoSize = true;
        lblLanguage.Location = new Point(12, 466);
        lblLanguage.Name = "lblLanguage";
        lblLanguage.Text = "Weergavetaal:";
        // 
        // cmbLanguage
        // 
        cmbLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbLanguage.Items.AddRange(new object[] { "Nederlands", "English" });
        cmbLanguage.Location = new Point(130, 462);
        cmbLanguage.Name = "cmbLanguage";
        cmbLanguage.Size = new Size(140, 23);
        cmbLanguage.TabIndex = 20;
        // 
        // btnOK
        // 
        btnOK.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnOK.Location = new Point(508, 496);
        btnOK.Name = "btnOK";
        btnOK.Size = new Size(80, 27);
        btnOK.Text = "Opslaan";
        btnOK.UseVisualStyleBackColor = true;
        btnOK.Click += btnOK_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.Location = new Point(598, 496);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(80, 27);
        btnCancel.Text = "Annuleren";
        btnCancel.UseVisualStyleBackColor = true;
        btnCancel.Click += btnCancel_Click;
        // 
        // SettingsForm
        // 
        AcceptButton = btnOK;
        CancelButton = btnCancel;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(690, 536);
        Controls.Add(btnCancel);
        Controls.Add(btnOK);
        Controls.Add(cmbLanguage);
        Controls.Add(lblLanguage);
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
        Controls.Add(chkUseIndex);
        Controls.Add(chkSearchAttachments);
        Controls.Add(chkSearchBody);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "SettingsForm";
        StartPosition = FormStartPosition.CenterParent;
        Text = "Instellingen";
        ClientSize = new Size(690, 536);
        ((System.ComponentModel.ISupportInitialize)nudSearchHistoryMax).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private CheckBox chkSearchBody;
    private CheckBox chkSearchAttachments;
    private CheckBox chkUseIndex;
    private Label lblMaxResults;
    private NumericUpDown nudMaxResults;
    private CheckBox chkExcludeAttachmentExt;
    private TextBox txtExcludedAttachmentExt;
    private Label lblStores;
    private Button btnRefreshStores;
    private CheckedListBox clbStores;
    private Label lblStoresStatus;
    private Button btnChooseExcludedFolders;
    private Label lblExcludedFolderSummary;
    private Button btnManageIndex;
    private Label lblLanguage;
    private ComboBox cmbLanguage;
    private Button btnOK;
    private Button btnCancel;
    private Label lblSearchHistoryMax;
    private NumericUpDown nudSearchHistoryMax;
    private Button btnClearHistory;
}
