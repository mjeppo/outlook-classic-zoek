namespace OutlookClassicSearch;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        lblQuery = new Label();
        txtQuery = new TextBox();
        chkSearchBody = new CheckBox();
        chkSearchAttachments = new CheckBox();
        chkUseDateRange = new CheckBox();
        dtpFrom = new DateTimePicker();
        dtpTo = new DateTimePicker();
        lblMaxResults = new Label();
        nudMaxResults = new NumericUpDown();
        chkUseIndex = new CheckBox();
        chkExcludeAttachmentExt = new CheckBox();
        txtExcludedAttachmentExt = new TextBox();
        lblStores = new Label();
        clbStores = new CheckedListBox();
        btnRefreshStores = new Button();
        btnChooseExcludedFolders = new Button();
        lblExcludedFolderSummary = new Label();
        btnManageIndex = new Button();
        btnSearch = new Button();
        btnCancel = new Button();
        progressBar = new ProgressBar();
        chkShowPreview = new CheckBox();
        lblFilterGlobal = new Label();
        txtResultFilter = new TextBox();
        lblFilterMailbox = new Label();
        txtFilterMailbox = new TextBox();
        lblFilterRecipients = new Label();
        txtFilterRecipients = new TextBox();
        dgvResults = new DataGridView();
        pnlPreview = new Panel();
        btnCopyPreview = new Button();
        txtPreview = new TextBox();
        statusStrip1 = new StatusStrip();
        toolStripStatusLabel1 = new ToolStripStatusLabel();
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).BeginInit();
        ((System.ComponentModel.ISupportInitialize)dgvResults).BeginInit();
        pnlPreview.SuspendLayout();
        statusStrip1.SuspendLayout();
        SuspendLayout();
        // 
        // lblQuery
        // 
        lblQuery.AutoSize = true;
        lblQuery.Location = new Point(12, 15);
        lblQuery.Name = "lblQuery";
        lblQuery.Size = new Size(63, 15);
        lblQuery.TabIndex = 0;
        lblQuery.Text = "Zoekterm:";
        // 
        // txtQuery
        // 
        txtQuery.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtQuery.Location = new Point(81, 12);
        txtQuery.Name = "txtQuery";
        txtQuery.PlaceholderText = "Bijv. factuur, ordernummer, klantnaam";
        txtQuery.Size = new Size(914, 23);
        txtQuery.TabIndex = 1;
        // 
        // chkSearchBody
        // 
        chkSearchBody.AutoSize = true;
        chkSearchBody.Checked = true;
        chkSearchBody.CheckState = CheckState.Checked;
        chkSearchBody.Location = new Point(81, 45);
        chkSearchBody.Name = "chkSearchBody";
        chkSearchBody.Size = new Size(122, 19);
        chkSearchBody.TabIndex = 2;
        chkSearchBody.Text = "Zoek ook in body";
        chkSearchBody.UseVisualStyleBackColor = true;
        // 
        // chkSearchAttachments
        // 
        chkSearchAttachments.AutoSize = true;
        chkSearchAttachments.Location = new Point(209, 45);
        chkSearchAttachments.Name = "chkSearchAttachments";
        chkSearchAttachments.Size = new Size(140, 19);
        chkSearchAttachments.TabIndex = 3;
        chkSearchAttachments.Text = "Zoek ook in bijlagen";
        chkSearchAttachments.UseVisualStyleBackColor = true;
        // 
        // chkUseDateRange
        // 
        chkUseDateRange.AutoSize = true;
        chkUseDateRange.Checked = true;
        chkUseDateRange.CheckState = CheckState.Checked;
        chkUseDateRange.Location = new Point(355, 45);
        chkUseDateRange.Name = "chkUseDateRange";
        chkUseDateRange.Size = new Size(103, 19);
        chkUseDateRange.TabIndex = 4;
        chkUseDateRange.Text = "Gebruik datum";
        chkUseDateRange.UseVisualStyleBackColor = true;
        // 
        // dtpFrom
        // 
        dtpFrom.Format = DateTimePickerFormat.Short;
        dtpFrom.Location = new Point(464, 43);
        dtpFrom.Name = "dtpFrom";
        dtpFrom.Size = new Size(94, 23);
        dtpFrom.TabIndex = 5;
        // 
        // dtpTo
        // 
        dtpTo.Format = DateTimePickerFormat.Short;
        dtpTo.Location = new Point(564, 43);
        dtpTo.Name = "dtpTo";
        dtpTo.Size = new Size(94, 23);
        dtpTo.TabIndex = 6;
        // 
        // lblMaxResults
        // 
        lblMaxResults.AutoSize = true;
        lblMaxResults.Location = new Point(664, 45);
        lblMaxResults.Name = "lblMaxResults";
        lblMaxResults.Size = new Size(73, 15);
        lblMaxResults.TabIndex = 7;
        lblMaxResults.Text = "Max regels:";
        // 
        // nudMaxResults
        // 
        nudMaxResults.Location = new Point(743, 43);
        nudMaxResults.Maximum = new decimal(new int[] { 5000, 0, 0, 0 });
        nudMaxResults.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
        nudMaxResults.Name = "nudMaxResults";
        nudMaxResults.Size = new Size(86, 23);
        nudMaxResults.TabIndex = 8;
        nudMaxResults.Value = new decimal(new int[] { 500, 0, 0, 0 });
        // 
        // chkUseIndex
        // 
        chkUseIndex.AutoSize = true;
        chkUseIndex.Checked = true;
        chkUseIndex.CheckState = CheckState.Checked;
        chkUseIndex.Location = new Point(835, 44);
        chkUseIndex.Name = "chkUseIndex";
        chkUseIndex.Size = new Size(160, 19);
        chkUseIndex.TabIndex = 9;
        chkUseIndex.Text = "Gebruik permanente index";
        chkUseIndex.UseVisualStyleBackColor = true;
        // 
        // chkExcludeAttachmentExt
        // 
        chkExcludeAttachmentExt.AutoSize = true;
        chkExcludeAttachmentExt.Checked = true;
        chkExcludeAttachmentExt.CheckState = CheckState.Checked;
        chkExcludeAttachmentExt.Location = new Point(12, 72);
        chkExcludeAttachmentExt.Name = "chkExcludeAttachmentExt";
        chkExcludeAttachmentExt.Size = new Size(187, 19);
        chkExcludeAttachmentExt.TabIndex = 10;
        chkExcludeAttachmentExt.Text = "Bijlage-extensies uitsluiten (;):";
        chkExcludeAttachmentExt.UseVisualStyleBackColor = true;
        // 
        // txtExcludedAttachmentExt
        // 
        txtExcludedAttachmentExt.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtExcludedAttachmentExt.Location = new Point(207, 71);
        txtExcludedAttachmentExt.Name = "txtExcludedAttachmentExt";
        txtExcludedAttachmentExt.Size = new Size(810, 23);
        txtExcludedAttachmentExt.TabIndex = 11;
        // 
        // lblStores
        // 
        lblStores.AutoSize = true;
        lblStores.Location = new Point(12, 105);
        lblStores.Name = "lblStores";
        lblStores.Size = new Size(127, 15);
        lblStores.TabIndex = 12;
        lblStores.Text = "Mailboxen (incl. shared):";
        // 
        // clbStores
        // 
        clbStores.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        clbStores.CheckOnClick = true;
        clbStores.FormattingEnabled = true;
        clbStores.Location = new Point(12, 123);
        clbStores.Name = "clbStores";
        clbStores.Size = new Size(983, 120);
        clbStores.TabIndex = 13;
        // 
        // btnRefreshStores
        // 
        btnRefreshStores.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnRefreshStores.Location = new Point(814, 100);
        btnRefreshStores.Name = "btnRefreshStores";
        btnRefreshStores.Size = new Size(181, 23);
        btnRefreshStores.TabIndex = 14;
        btnRefreshStores.Text = "Mailboxen vernieuwen";
        btnRefreshStores.UseVisualStyleBackColor = true;
        btnRefreshStores.Click += btnRefreshStores_Click;
        // 
        // btnChooseExcludedFolders
        // 
        btnChooseExcludedFolders.Location = new Point(12, 253);
        btnChooseExcludedFolders.Name = "btnChooseExcludedFolders";
        btnChooseExcludedFolders.Size = new Size(220, 27);
        btnChooseExcludedFolders.TabIndex = 15;
        btnChooseExcludedFolders.Text = "Mappen uitsluiten van zoeken";
        btnChooseExcludedFolders.UseVisualStyleBackColor = true;
        btnChooseExcludedFolders.Click += btnChooseExcludedFolders_Click;
        // 
        // lblExcludedFolderSummary
        // 
        lblExcludedFolderSummary.AutoSize = true;
        lblExcludedFolderSummary.Location = new Point(238, 259);
        lblExcludedFolderSummary.Name = "lblExcludedFolderSummary";
        lblExcludedFolderSummary.Size = new Size(123, 15);
        lblExcludedFolderSummary.TabIndex = 16;
        lblExcludedFolderSummary.Text = "Uitgesloten mappen: 0";
        // 
        // btnManageIndex
        // 
        btnManageIndex.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnManageIndex.Location = new Point(814, 253);
        btnManageIndex.Name = "btnManageIndex";
        btnManageIndex.Size = new Size(181, 27);
        btnManageIndex.TabIndex = 17;
        btnManageIndex.Text = "Index beheren";
        btnManageIndex.UseVisualStyleBackColor = true;
        btnManageIndex.Click += btnManageIndex_Click;
        // 
        // btnSearch
        // 
        btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnSearch.Location = new Point(814, 286);
        btnSearch.Name = "btnSearch";
        btnSearch.Size = new Size(86, 27);
        btnSearch.TabIndex = 18;
        btnSearch.Text = "Zoeken";
        btnSearch.UseVisualStyleBackColor = true;
        btnSearch.Click += btnSearch_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnCancel.Enabled = false;
        btnCancel.Location = new Point(909, 286);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(86, 27);
        btnCancel.TabIndex = 19;
        btnCancel.Text = "Annuleren";
        btnCancel.UseVisualStyleBackColor = true;
        btnCancel.Click += btnCancel_Click;
        // 
        // progressBar
        // 
        progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        progressBar.Location = new Point(632, 290);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(176, 18);
        progressBar.Style = ProgressBarStyle.Marquee;
        progressBar.TabIndex = 20;
        progressBar.Visible = false;
        // 
        // chkShowPreview
        // 
        chkShowPreview.AutoSize = true;
        chkShowPreview.Location = new Point(12, 324);
        chkShowPreview.Name = "chkShowPreview";
        chkShowPreview.Size = new Size(139, 19);
        chkShowPreview.TabIndex = 21;
        chkShowPreview.Text = "Toon voorbeeldvenster";
        chkShowPreview.UseVisualStyleBackColor = true;
        // 
        // lblFilterGlobal
        // 
        lblFilterGlobal.AutoSize = true;
        lblFilterGlobal.Location = new Point(169, 325);
        lblFilterGlobal.Name = "lblFilterGlobal";
        lblFilterGlobal.Size = new Size(46, 15);
        lblFilterGlobal.TabIndex = 22;
        lblFilterGlobal.Text = "Algem.:";
        // 
        // txtResultFilter
        // 
        txtResultFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        txtResultFilter.Location = new Point(221, 321);
        txtResultFilter.Name = "txtResultFilter";
        txtResultFilter.PlaceholderText = "Alle kolommen";
        txtResultFilter.Size = new Size(190, 23);
        txtResultFilter.TabIndex = 23;
        // 
        // lblFilterMailbox
        // 
        lblFilterMailbox.AutoSize = true;
        lblFilterMailbox.Location = new Point(417, 325);
        lblFilterMailbox.Name = "lblFilterMailbox";
        lblFilterMailbox.Size = new Size(52, 15);
        lblFilterMailbox.TabIndex = 24;
        lblFilterMailbox.Text = "Mailbox:";
        // 
        // txtFilterMailbox
        // 
        txtFilterMailbox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        txtFilterMailbox.Location = new Point(475, 321);
        txtFilterMailbox.Name = "txtFilterMailbox";
        txtFilterMailbox.PlaceholderText = "Alleen mailbox";
        txtFilterMailbox.Size = new Size(190, 23);
        txtFilterMailbox.TabIndex = 25;
        // 
        // lblFilterRecipients
        // 
        lblFilterRecipients.AutoSize = true;
        lblFilterRecipients.Location = new Point(671, 325);
        lblFilterRecipients.Name = "lblFilterRecipients";
        lblFilterRecipients.Size = new Size(81, 15);
        lblFilterRecipients.TabIndex = 26;
        lblFilterRecipients.Text = "Geadresseerde:";
        // 
        // txtFilterRecipients
        // 
        txtFilterRecipients.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtFilterRecipients.Location = new Point(758, 321);
        txtFilterRecipients.Name = "txtFilterRecipients";
        txtFilterRecipients.PlaceholderText = "Alleen geadresseerde";
        txtFilterRecipients.Size = new Size(237, 23);
        txtFilterRecipients.TabIndex = 27;
        // 
        // dgvResults
        // 
        dgvResults.AllowUserToAddRows = false;
        dgvResults.AllowUserToDeleteRows = false;
        dgvResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        dgvResults.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResults.Location = new Point(12, 352);
        dgvResults.Name = "dgvResults";
        dgvResults.ReadOnly = true;
        dgvResults.RowHeadersVisible = false;
        dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvResults.Size = new Size(983, 250);
        dgvResults.TabIndex = 28;
        dgvResults.CellDoubleClick += dgvResults_CellDoubleClick;
        // 
        // pnlPreview
        // 
        pnlPreview.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        pnlPreview.BorderStyle = BorderStyle.FixedSingle;
        pnlPreview.Controls.Add(btnCopyPreview);
        pnlPreview.Controls.Add(txtPreview);
        pnlPreview.Location = new Point(12, 608);
        pnlPreview.Name = "pnlPreview";
        pnlPreview.Size = new Size(983, 61);
        pnlPreview.TabIndex = 29;
        pnlPreview.Visible = false;
        // 
        // btnCopyPreview
        // 
        btnCopyPreview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnCopyPreview.Location = new Point(858, 6);
        btnCopyPreview.Name = "btnCopyPreview";
        btnCopyPreview.Size = new Size(115, 24);
        btnCopyPreview.TabIndex = 1;
        btnCopyPreview.Text = "Kopieer voorbeeld";
        btnCopyPreview.UseVisualStyleBackColor = true;
        btnCopyPreview.Click += btnCopyPreview_Click;
        // 
        // txtPreview
        // 
        txtPreview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        txtPreview.Location = new Point(8, 36);
        txtPreview.Multiline = true;
        txtPreview.Name = "txtPreview";
        txtPreview.ReadOnly = true;
        txtPreview.ScrollBars = ScrollBars.Both;
        txtPreview.Size = new Size(965, 18);
        txtPreview.TabIndex = 0;
        // 
        // statusStrip1
        // 
        statusStrip1.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel1 });
        statusStrip1.Dock = DockStyle.Bottom;
        statusStrip1.Name = "statusStrip1";
        statusStrip1.Size = new Size(1007, 22);
        statusStrip1.TabIndex = 22;
        statusStrip1.Text = "statusStrip1";
        // 
        // toolStripStatusLabel1
        // 
        toolStripStatusLabel1.Name = "toolStripStatusLabel1";
        toolStripStatusLabel1.Size = new Size(126, 17);
        toolStripStatusLabel1.Text = "Klaar voor zoekopdracht";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1007, 697);
        Controls.Add(statusStrip1);
        Controls.Add(pnlPreview);
        Controls.Add(dgvResults);
        Controls.Add(txtFilterRecipients);
        Controls.Add(lblFilterRecipients);
        Controls.Add(txtFilterMailbox);
        Controls.Add(lblFilterMailbox);
        Controls.Add(txtResultFilter);
        Controls.Add(lblFilterGlobal);
        Controls.Add(chkShowPreview);
        Controls.Add(progressBar);
        Controls.Add(btnCancel);
        Controls.Add(btnSearch);
        Controls.Add(btnManageIndex);
        Controls.Add(lblExcludedFolderSummary);
        Controls.Add(btnChooseExcludedFolders);
        Controls.Add(btnRefreshStores);
        Controls.Add(clbStores);
        Controls.Add(lblStores);
        Controls.Add(txtExcludedAttachmentExt);
        Controls.Add(chkExcludeAttachmentExt);
        Controls.Add(chkUseIndex);
        Controls.Add(nudMaxResults);
        Controls.Add(lblMaxResults);
        Controls.Add(dtpTo);
        Controls.Add(dtpFrom);
        Controls.Add(chkUseDateRange);
        Controls.Add(chkSearchAttachments);
        Controls.Add(chkSearchBody);
        Controls.Add(txtQuery);
        Controls.Add(lblQuery);
        MinimumSize = new Size(920, 560);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Outlook Classic Search";
        Load += Form1_Load;
        FormClosing += Form1_FormClosing;
        ((System.ComponentModel.ISupportInitialize)nudMaxResults).EndInit();
        ((System.ComponentModel.ISupportInitialize)dgvResults).EndInit();
        pnlPreview.ResumeLayout(false);
        pnlPreview.PerformLayout();
        statusStrip1.ResumeLayout(false);
        statusStrip1.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Label lblQuery;
    private TextBox txtQuery;
    private CheckBox chkSearchBody;
    private CheckBox chkSearchAttachments;
    private CheckBox chkUseDateRange;
    private DateTimePicker dtpFrom;
    private DateTimePicker dtpTo;
    private Label lblMaxResults;
    private NumericUpDown nudMaxResults;
    private CheckBox chkUseIndex;
    private CheckBox chkExcludeAttachmentExt;
    private TextBox txtExcludedAttachmentExt;
    private Label lblStores;
    private CheckedListBox clbStores;
    private Button btnRefreshStores;
    private Button btnChooseExcludedFolders;
    private Label lblExcludedFolderSummary;
    private Button btnManageIndex;
    private Button btnSearch;
    private Button btnCancel;
    private ProgressBar progressBar;
    private CheckBox chkShowPreview;
    private Label lblFilterGlobal;
    private TextBox txtResultFilter;
    private Label lblFilterMailbox;
    private TextBox txtFilterMailbox;
    private Label lblFilterRecipients;
    private TextBox txtFilterRecipients;
    private DataGridView dgvResults;
    private Panel pnlPreview;
    private Button btnCopyPreview;
    private TextBox txtPreview;
    private StatusStrip statusStrip1;
    private ToolStripStatusLabel toolStripStatusLabel1;
}
