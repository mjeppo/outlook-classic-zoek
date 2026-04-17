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
        menuStrip1 = new MenuStrip();
        mnuInstellingen = new ToolStripMenuItem();
        mnuHelp = new ToolStripMenuItem();
        mnuHelpHelp = new ToolStripMenuItem();
        mnuHelpInfo = new ToolStripMenuItem();
        lblQuery = new Label();
        cmbQuery = new ComboBox();
        btnSearch = new Button();
        btnCancel = new Button();
        lblShowPreview = new Label();
        chkShowPreview = new MaterialSkin.Controls.MaterialSwitch();
        lblUseDateRange = new Label();
        chkUseDateRange = new MaterialSkin.Controls.MaterialSwitch();
        dtpFrom = new DateTimePicker();
        dtpTo = new DateTimePicker();
        pnlFilters = new Panel();
        btnColMailbox = new Button();
        btnColFolderPath = new Button();
        txtColDate = new BorderedTextBox();
        txtColSubject = new BorderedTextBox();
        btnColSender = new Button();
        btnColRecipients = new Button();
        btnColHasAttachment = new Button();
        chkHasAttachment = new MaterialSkin.Controls.MaterialSwitch();
        splitContainer = new SplitContainer();
        dgvResults = new DataGridView();
        pnlPreview = new Panel();
        txtPreview = new TextBox();
        pnlPreviewTop = new Panel();
        btnClosePreview = new Button();
        btnCopyPreview = new Button();
        statusStrip1 = new StatusStrip();
        toolStripProgressBar = new ToolStripProgressBar();
        toolStripStatusLabel1 = new ToolStripStatusLabel();
        chkSearchBody = new MaterialSkin.Controls.MaterialSwitch();
        chkSearchAttachments = new MaterialSkin.Controls.MaterialSwitch();
        lblSearchBody = new Label();
        lblSearchAttachments = new Label();
        lblMainMaxResults = new Label();
        nudMainMaxResults = new NumericUpDown();
        btnExcludeFolders = new Button();
        lblExcludedFolderSummary = new Label();
        menuStrip1.SuspendLayout();
        pnlFilters.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer).BeginInit();
        splitContainer.Panel1.SuspendLayout();
        splitContainer.Panel2.SuspendLayout();
        splitContainer.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResults).BeginInit();
        pnlPreview.SuspendLayout();
        pnlPreviewTop.SuspendLayout();
        statusStrip1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudMainMaxResults).BeginInit();
        SuspendLayout();
        // 
        // menuStrip1
        // 
        menuStrip1.Items.AddRange(new ToolStripItem[] { mnuInstellingen, mnuHelp });
        menuStrip1.Location = new Point(3, 64);
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(1094, 24);
        menuStrip1.TabIndex = 0;
        // 
        // mnuInstellingen
        // 
        mnuInstellingen.Name = "mnuInstellingen";
        mnuInstellingen.Size = new Size(80, 20);
        mnuInstellingen.Text = "Instellingen";
        mnuInstellingen.Click += mnuInstellingen_Click;
        // 
        // mnuHelp
        // 
        mnuHelp.DropDownItems.AddRange(new ToolStripItem[] { mnuHelpHelp, mnuHelpInfo });
        mnuHelp.Name = "mnuHelp";
        mnuHelp.Size = new Size(44, 20);
        mnuHelp.Text = "Help";
        // 
        // mnuHelpHelp
        // 
        mnuHelpHelp.Name = "mnuHelpHelp";
        mnuHelpHelp.Size = new Size(99, 22);
        mnuHelpHelp.Text = "Help";
        mnuHelpHelp.Click += mnuHelpHelp_Click;
        // 
        // mnuHelpInfo
        // 
        mnuHelpInfo.Name = "mnuHelpInfo";
        mnuHelpInfo.Size = new Size(99, 22);
        mnuHelpInfo.Text = "Info";
        mnuHelpInfo.Click += mnuHelpInfo_Click;
        // 
        // lblQuery
        // 
        lblQuery.AutoSize = true;
        lblQuery.Location = new Point(8, 96);
        lblQuery.Name = "lblQuery";
        lblQuery.Size = new Size(61, 15);
        lblQuery.TabIndex = 14;
        lblQuery.Text = "Zoekterm:";
        // 
        // cmbQuery
        // 
        cmbQuery.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        cmbQuery.Location = new Point(75, 92);
        cmbQuery.Name = "cmbQuery";
        cmbQuery.Size = new Size(829, 23);
        cmbQuery.TabIndex = 1;
        cmbQuery.KeyDown += cmbQuery_KeyDown;
        // 
        // btnSearch
        // 
        btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnSearch.Location = new Point(912, 91);
        btnSearch.Name = "btnSearch";
        btnSearch.Size = new Size(86, 27);
        btnSearch.TabIndex = 2;
        btnSearch.Text = "Zoeken";
        btnSearch.UseVisualStyleBackColor = true;
        btnSearch.Click += btnSearch_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnCancel.Enabled = false;
        btnCancel.Location = new Point(1006, 91);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(86, 27);
        btnCancel.TabIndex = 3;
        btnCancel.Text = "Annuleren";
        btnCancel.UseVisualStyleBackColor = true;
        btnCancel.Click += btnCancel_Click;
        // 
        // lblShowPreview
        // 
        lblShowPreview.AutoSize = true;
        lblShowPreview.Cursor = Cursors.Hand;
        lblShowPreview.Location = new Point(59, 134);
        lblShowPreview.Name = "lblShowPreview";
        lblShowPreview.Size = new Size(128, 15);
        lblShowPreview.TabIndex = 100;
        lblShowPreview.Text = "Toon voorbeeldvenster";
        lblShowPreview.Click += lblShowPreview_Click;
        // 
        // chkShowPreview
        // 
        chkShowPreview.AutoSize = true;
        chkShowPreview.Depth = 0;
        chkShowPreview.Location = new Point(8, 124);
        chkShowPreview.Margin = new Padding(0);
        chkShowPreview.MouseLocation = new Point(-1, -1);
        chkShowPreview.MouseState = MaterialSkin.MouseState.HOVER;
        chkShowPreview.Name = "chkShowPreview";
        chkShowPreview.Ripple = true;
        chkShowPreview.Size = new Size(58, 37);
        chkShowPreview.TabIndex = 13;
        chkShowPreview.UseVisualStyleBackColor = true;
        chkShowPreview.CheckedChanged += chkShowPreview_CheckedChanged;
        // 
        // lblUseDateRange
        // 
        lblUseDateRange.AutoSize = true;
        lblUseDateRange.Cursor = Cursors.Hand;
        lblUseDateRange.Location = new Point(243, 134);
        lblUseDateRange.Name = "lblUseDateRange";
        lblUseDateRange.Size = new Size(86, 15);
        lblUseDateRange.TabIndex = 101;
        lblUseDateRange.Text = "Gebruik datum";
        lblUseDateRange.Click += lblUseDateRange_Click;
        // 
        // chkUseDateRange
        // 
        chkUseDateRange.AutoSize = true;
        chkUseDateRange.Checked = true;
        chkUseDateRange.CheckState = CheckState.Checked;
        chkUseDateRange.Depth = 0;
        chkUseDateRange.Location = new Point(190, 124);
        chkUseDateRange.Margin = new Padding(0);
        chkUseDateRange.MouseLocation = new Point(-1, -1);
        chkUseDateRange.MouseState = MaterialSkin.MouseState.HOVER;
        chkUseDateRange.Name = "chkUseDateRange";
        chkUseDateRange.Ripple = true;
        chkUseDateRange.Size = new Size(58, 37);
        chkUseDateRange.TabIndex = 12;
        chkUseDateRange.UseVisualStyleBackColor = true;
        // 
        // dtpFrom
        // 
        dtpFrom.CustomFormat = "dd-MM-yyyy";
        dtpFrom.Format = DateTimePickerFormat.Custom;
        dtpFrom.Location = new Point(333, 129);
        dtpFrom.Name = "dtpFrom";
        dtpFrom.Size = new Size(94, 23);
        dtpFrom.TabIndex = 5;
        // 
        // dtpTo
        // 
        dtpTo.CustomFormat = "dd-MM-yyyy";
        dtpTo.Format = DateTimePickerFormat.Custom;
        dtpTo.Location = new Point(433, 129);
        dtpTo.Name = "dtpTo";
        dtpTo.Size = new Size(94, 23);
        dtpTo.TabIndex = 6;
        // 
        // pnlFilters
        // 
        pnlFilters.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        pnlFilters.Controls.Add(btnColMailbox);
        pnlFilters.Controls.Add(btnColFolderPath);
        pnlFilters.Controls.Add(txtColDate);
        pnlFilters.Controls.Add(txtColSubject);
        pnlFilters.Controls.Add(btnColSender);
        pnlFilters.Controls.Add(btnColRecipients);
        pnlFilters.Controls.Add(btnColHasAttachment);
        pnlFilters.Location = new Point(8, 190);
        pnlFilters.Name = "pnlFilters";
        pnlFilters.Size = new Size(1084, 27);
        pnlFilters.TabIndex = 7;
        // 
        // btnColMailbox
        // 
        btnColMailbox.FlatStyle = FlatStyle.Flat;
        btnColMailbox.Location = new Point(0, 0);
        btnColMailbox.Name = "btnColMailbox";
        btnColMailbox.Size = new Size(75, 23);
        btnColMailbox.TabIndex = 0;
        btnColMailbox.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // btnColFolderPath
        // 
        btnColFolderPath.FlatStyle = FlatStyle.Flat;
        btnColFolderPath.Location = new Point(0, 0);
        btnColFolderPath.Name = "btnColFolderPath";
        btnColFolderPath.Size = new Size(75, 23);
        btnColFolderPath.TabIndex = 1;
        btnColFolderPath.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // txtColDate
        // 
        txtColDate.BorderStyle = BorderStyle.FixedSingle;
        txtColDate.Font = new Font("Segoe UI", 9.75F);
        txtColDate.Location = new Point(0, 0);
        txtColDate.Multiline = true;
        txtColDate.Name = "txtColDate";
        txtColDate.PlaceholderText = "Filter: Datum";
        txtColDate.Size = new Size(100, 27);
        txtColDate.TabIndex = 2;
        // 
        // txtColSubject
        // 
        txtColSubject.BorderStyle = BorderStyle.FixedSingle;
        txtColSubject.Font = new Font("Segoe UI", 9.75F);
        txtColSubject.Location = new Point(0, 0);
        txtColSubject.Multiline = true;
        txtColSubject.Name = "txtColSubject";
        txtColSubject.PlaceholderText = "Filter: Onderwerp";
        txtColSubject.Size = new Size(100, 27);
        txtColSubject.TabIndex = 3;
        // 
        // btnColSender
        // 
        btnColSender.FlatStyle = FlatStyle.Flat;
        btnColSender.Location = new Point(0, 0);
        btnColSender.Name = "btnColSender";
        btnColSender.Size = new Size(75, 23);
        btnColSender.TabIndex = 4;
        btnColSender.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // btnColRecipients
        // 
        btnColRecipients.FlatStyle = FlatStyle.Flat;
        btnColRecipients.Location = new Point(0, 0);
        btnColRecipients.Name = "btnColRecipients";
        btnColRecipients.Size = new Size(75, 23);
        btnColRecipients.TabIndex = 5;
        btnColRecipients.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // btnColHasAttachment
        // 
        btnColHasAttachment.FlatStyle = FlatStyle.Flat;
        btnColHasAttachment.Location = new Point(0, 0);
        btnColHasAttachment.Name = "btnColHasAttachment";
        btnColHasAttachment.Size = new Size(75, 23);
        btnColHasAttachment.TabIndex = 6;
        btnColHasAttachment.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // chkHasAttachment
        // 
        chkHasAttachment.Depth = 0;
        chkHasAttachment.Location = new Point(0, 0);
        chkHasAttachment.Margin = new Padding(0);
        chkHasAttachment.MouseLocation = new Point(-1, -1);
        chkHasAttachment.MouseState = MaterialSkin.MouseState.HOVER;
        chkHasAttachment.Name = "chkHasAttachment";
        chkHasAttachment.Ripple = true;
        chkHasAttachment.Size = new Size(104, 24);
        chkHasAttachment.TabIndex = 0;
        // 
        // splitContainer
        // 
        splitContainer.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        splitContainer.Location = new Point(8, 222);
        splitContainer.Name = "splitContainer";
        // 
        // splitContainer.Panel1
        // 
        splitContainer.Panel1.Controls.Add(dgvResults);
        splitContainer.Panel1MinSize = 200;
        // 
        // splitContainer.Panel2
        // 
        splitContainer.Panel2.Controls.Add(pnlPreview);
        splitContainer.Panel2Collapsed = true;
        splitContainer.Panel2MinSize = 150;
        splitContainer.Size = new Size(1084, 470);
        splitContainer.SplitterDistance = 200;
        splitContainer.SplitterWidth = 6;
        splitContainer.TabIndex = 8;
        // 
        // dgvResults
        // 
        dgvResults.AllowUserToAddRows = false;
        dgvResults.AllowUserToDeleteRows = false;
        dgvResults.AllowUserToOrderColumns = true;
        dgvResults.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResults.Dock = DockStyle.Fill;
        dgvResults.Location = new Point(0, 0);
        dgvResults.Name = "dgvResults";
        dgvResults.ReadOnly = true;
        dgvResults.RowHeadersVisible = false;
        dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvResults.Size = new Size(1084, 470);
        dgvResults.TabIndex = 0;
        dgvResults.CellDoubleClick += dgvResults_CellDoubleClick;
        // 
        // pnlPreview
        // 
        pnlPreview.BorderStyle = BorderStyle.FixedSingle;
        pnlPreview.Controls.Add(txtPreview);
        pnlPreview.Controls.Add(pnlPreviewTop);
        pnlPreview.Dock = DockStyle.Fill;
        pnlPreview.Location = new Point(0, 0);
        pnlPreview.Name = "pnlPreview";
        pnlPreview.Size = new Size(96, 100);
        pnlPreview.TabIndex = 0;
        // 
        // txtPreview
        // 
        txtPreview.Dock = DockStyle.Fill;
        txtPreview.Location = new Point(0, 34);
        txtPreview.Multiline = true;
        txtPreview.Name = "txtPreview";
        txtPreview.ReadOnly = true;
        txtPreview.ScrollBars = ScrollBars.Both;
        txtPreview.Size = new Size(94, 64);
        txtPreview.TabIndex = 1;
        // 
        // pnlPreviewTop
        // 
        pnlPreviewTop.Controls.Add(btnClosePreview);
        pnlPreviewTop.Controls.Add(btnCopyPreview);
        pnlPreviewTop.Dock = DockStyle.Top;
        pnlPreviewTop.Location = new Point(0, 0);
        pnlPreviewTop.Name = "pnlPreviewTop";
        pnlPreviewTop.Size = new Size(94, 34);
        pnlPreviewTop.TabIndex = 0;
        // 
        // btnClosePreview
        // 
        btnClosePreview.Location = new Point(6, 5);
        btnClosePreview.Name = "btnClosePreview";
        btnClosePreview.Size = new Size(120, 24);
        btnClosePreview.TabIndex = 0;
        btnClosePreview.Text = "Venster sluiten";
        btnClosePreview.UseVisualStyleBackColor = true;
        btnClosePreview.Click += btnClosePreview_Click;
        // 
        // btnCopyPreview
        // 
        btnCopyPreview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnCopyPreview.Location = new Point(-27, 5);
        btnCopyPreview.Name = "btnCopyPreview";
        btnCopyPreview.Size = new Size(115, 24);
        btnCopyPreview.TabIndex = 1;
        btnCopyPreview.Text = "Kopieer voorbeeld";
        btnCopyPreview.UseVisualStyleBackColor = true;
        btnCopyPreview.Click += btnCopyPreview_Click;
        // 
        // statusStrip1
        // 
        statusStrip1.Items.AddRange(new ToolStripItem[] { toolStripProgressBar, toolStripStatusLabel1 });
        statusStrip1.Location = new Point(3, 693);
        statusStrip1.Name = "statusStrip1";
        statusStrip1.Size = new Size(1094, 22);
        statusStrip1.TabIndex = 1;
        // 
        // toolStripProgressBar
        // 
        toolStripProgressBar.Name = "toolStripProgressBar";
        toolStripProgressBar.Size = new Size(120, 16);
        toolStripProgressBar.Style = ProgressBarStyle.Marquee;
        toolStripProgressBar.Visible = false;
        // 
        // toolStripStatusLabel1
        // 
        toolStripStatusLabel1.Name = "toolStripStatusLabel1";
        toolStripStatusLabel1.Size = new Size(1079, 17);
        toolStripStatusLabel1.Spring = true;
        toolStripStatusLabel1.Text = "Klaar voor zoekopdracht";
        toolStripStatusLabel1.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // chkSearchBody
        // 
        chkSearchBody.AutoSize = true;
        chkSearchBody.Depth = 0;
        chkSearchBody.Location = new Point(8, 155);
        chkSearchBody.Margin = new Padding(0);
        chkSearchBody.MouseLocation = new Point(-1, -1);
        chkSearchBody.MouseState = MaterialSkin.MouseState.HOVER;
        chkSearchBody.Name = "chkSearchBody";
        chkSearchBody.Ripple = true;
        chkSearchBody.Size = new Size(58, 37);
        chkSearchBody.TabIndex = 7;
        chkSearchBody.UseVisualStyleBackColor = true;
        // 
        // chkSearchAttachments
        // 
        chkSearchAttachments.AutoSize = true;
        chkSearchAttachments.Depth = 0;
        chkSearchAttachments.Location = new Point(190, 155);
        chkSearchAttachments.Margin = new Padding(0);
        chkSearchAttachments.MouseLocation = new Point(-1, -1);
        chkSearchAttachments.MouseState = MaterialSkin.MouseState.HOVER;
        chkSearchAttachments.Name = "chkSearchAttachments";
        chkSearchAttachments.Ripple = true;
        chkSearchAttachments.Size = new Size(58, 37);
        chkSearchAttachments.TabIndex = 8;
        chkSearchAttachments.UseVisualStyleBackColor = true;
        // 
        // lblSearchBody
        // 
        lblSearchBody.AutoSize = true;
        lblSearchBody.Cursor = Cursors.Hand;
        lblSearchBody.Location = new Point(59, 164);
        lblSearchBody.Name = "lblSearchBody";
        lblSearchBody.Size = new Size(99, 15);
        lblSearchBody.TabIndex = 102;
        lblSearchBody.Text = "Zoek ook in body";
        lblSearchBody.Click += lblSearchBody_Click;
        // 
        // lblSearchAttachments
        // 
        lblSearchAttachments.AutoSize = true;
        lblSearchAttachments.Cursor = Cursors.Hand;
        lblSearchAttachments.Location = new Point(243, 165);
        lblSearchAttachments.Name = "lblSearchAttachments";
        lblSearchAttachments.Size = new Size(114, 15);
        lblSearchAttachments.TabIndex = 103;
        lblSearchAttachments.Text = "Zoek ook in bijlagen";
        lblSearchAttachments.Click += lblSearchAttachments_Click;
        // 
        // lblMainMaxResults
        // 
        lblMainMaxResults.AutoSize = true;
        lblMainMaxResults.Location = new Point(368, 165);
        lblMainMaxResults.Name = "lblMainMaxResults";
        lblMainMaxResults.Size = new Size(87, 15);
        lblMainMaxResults.TabIndex = 11;
        lblMainMaxResults.Text = "Max resultaten:";
        // 
        // nudMainMaxResults
        // 
        nudMainMaxResults.Location = new Point(454, 163);
        nudMainMaxResults.Maximum = new decimal(new int[] { 5000, 0, 0, 0 });
        nudMainMaxResults.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
        nudMainMaxResults.Name = "nudMainMaxResults";
        nudMainMaxResults.Size = new Size(72, 23);
        nudMainMaxResults.TabIndex = 9;
        nudMainMaxResults.Value = new decimal(new int[] { 500, 0, 0, 0 });
        // 
        // btnExcludeFolders
        // 
        btnExcludeFolders.Location = new Point(542, 127);
        btnExcludeFolders.Name = "btnExcludeFolders";
        btnExcludeFolders.Size = new Size(188, 27);
        btnExcludeFolders.TabIndex = 10;
        btnExcludeFolders.Text = "Mappen uitsluiten van zoeken...";
        btnExcludeFolders.UseVisualStyleBackColor = true;
        btnExcludeFolders.Click += btnExcludeFolders_Click;
        // 
        // lblExcludedFolderSummary
        // 
        lblExcludedFolderSummary.AutoSize = true;
        lblExcludedFolderSummary.Location = new Point(730, 133);
        lblExcludedFolderSummary.Name = "lblExcludedFolderSummary";
        lblExcludedFolderSummary.Size = new Size(126, 15);
        lblExcludedFolderSummary.TabIndex = 9;
        lblExcludedFolderSummary.Text = "Uitgesloten mappen: 0";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1100, 718);
        Controls.Add(menuStrip1);
        Controls.Add(statusStrip1);
        Controls.Add(splitContainer);
        Controls.Add(pnlFilters);
        Controls.Add(lblExcludedFolderSummary);
        Controls.Add(btnExcludeFolders);
        Controls.Add(nudMainMaxResults);
        Controls.Add(lblMainMaxResults);
        Controls.Add(lblSearchAttachments);
        Controls.Add(chkSearchAttachments);
        Controls.Add(lblSearchBody);
        Controls.Add(chkSearchBody);
        Controls.Add(dtpTo);
        Controls.Add(dtpFrom);
        Controls.Add(lblUseDateRange);
        Controls.Add(chkUseDateRange);
        Controls.Add(lblShowPreview);
        Controls.Add(chkShowPreview);
        Controls.Add(btnCancel);
        Controls.Add(btnSearch);
        Controls.Add(cmbQuery);
        Controls.Add(lblQuery);
        MainMenuStrip = menuStrip1;
        MinimumSize = new Size(800, 500);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Outlook Classic Search";
        FormClosing += Form1_FormClosing;
        Load += Form1_Load;
        menuStrip1.ResumeLayout(false);
        menuStrip1.PerformLayout();
        pnlFilters.ResumeLayout(false);
        pnlFilters.PerformLayout();
        splitContainer.Panel1.ResumeLayout(false);
        splitContainer.Panel2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)splitContainer).EndInit();
        splitContainer.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResults).EndInit();
        pnlPreview.ResumeLayout(false);
        pnlPreview.PerformLayout();
        pnlPreviewTop.ResumeLayout(false);
        statusStrip1.ResumeLayout(false);
        statusStrip1.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)nudMainMaxResults).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private MenuStrip menuStrip1;
    private ToolStripMenuItem mnuInstellingen;
    private ToolStripMenuItem mnuHelp;
    private ToolStripMenuItem mnuHelpHelp;
    private ToolStripMenuItem mnuHelpInfo;
    private Label lblQuery;
    private ComboBox cmbQuery;
    private Button btnSearch;
    private Button btnCancel;
    private Label lblShowPreview;
    private MaterialSkin.Controls.MaterialSwitch chkShowPreview;
    private Label lblUseDateRange;
    private MaterialSkin.Controls.MaterialSwitch chkUseDateRange;
    private DateTimePicker dtpFrom;
    private DateTimePicker dtpTo;
    private Panel pnlFilters;
    private Button btnColMailbox;
    private Button btnColFolderPath;
    private BorderedTextBox txtColDate;
    private BorderedTextBox txtColSubject;
    private Button btnColSender;
    private Button btnColRecipients;
    private Button btnColHasAttachment;
    private MaterialSkin.Controls.MaterialSwitch chkHasAttachment; // kept for designer compatibility but unused in UI
    private SplitContainer splitContainer;
    private Panel pnlPreviewTop;
    private DataGridView dgvResults;
    private Panel pnlPreview;
    private Button btnClosePreview;
    private Button btnCopyPreview;
    private TextBox txtPreview;
    private StatusStrip statusStrip1;
    private ToolStripStatusLabel toolStripStatusLabel1;
    private ToolStripProgressBar toolStripProgressBar;
    private Label lblSearchBody;
    private MaterialSkin.Controls.MaterialSwitch chkSearchBody;
    private Label lblSearchAttachments;
    private MaterialSkin.Controls.MaterialSwitch chkSearchAttachments;
    private Label lblMainMaxResults;
    private NumericUpDown nudMainMaxResults;
    private Button btnExcludeFolders;
    private Label lblExcludedFolderSummary;
}
