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
        menuStrip1 = new MenuStrip();
        mnuInstellingen = new ToolStripMenuItem();
        mnuHelp = new ToolStripMenuItem();
        mnuHelpHelp = new ToolStripMenuItem();
        mnuHelpInfo = new ToolStripMenuItem();
        lblQuery = new Label();
        cmbQuery = new ComboBox();
        btnSearch = new Button();
        btnCancel = new Button();
        chkShowPreview = new CheckBox();
        chkUseDateRange = new CheckBox();
        dtpFrom = new DateTimePicker();
        dtpTo = new DateTimePicker();
        pnlFilters = new Panel();
        txtColMailbox = new TextBox();
        txtColFolderPath = new TextBox();
        txtColDate = new TextBox();
        txtColSubject = new TextBox();
        txtColSender = new TextBox();
        txtColRecipients = new TextBox();
        splitContainer = new SplitContainer();
        dgvResults = new DataGridView();
        pnlPreview = new Panel();
        pnlPreviewTop = new Panel();
        btnClosePreview = new Button();
        btnCopyPreview = new Button();
        txtPreview = new TextBox();
        statusStrip1 = new StatusStrip();
        toolStripStatusLabel1 = new ToolStripStatusLabel();
        toolStripProgressBar = new ToolStripProgressBar();
        menuStrip1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer).BeginInit();
        splitContainer.Panel1.SuspendLayout();
        splitContainer.Panel2.SuspendLayout();
        splitContainer.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResults).BeginInit();
        pnlPreviewTop.SuspendLayout();
        pnlPreview.SuspendLayout();
        statusStrip1.SuspendLayout();
        SuspendLayout();
        // 
        // menuStrip1
        // 
        menuStrip1.Items.AddRange(new ToolStripItem[] { mnuInstellingen, mnuHelp });
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(1100, 24);
        // 
        // mnuInstellingen
        // 
        mnuInstellingen.Name = "mnuInstellingen";
        mnuInstellingen.Text = "⚙ Instellingen";
        mnuInstellingen.Click += mnuInstellingen_Click;
        // 
        // mnuHelp
        // 
        mnuHelp.DropDownItems.AddRange(new ToolStripItem[] { mnuHelpHelp, mnuHelpInfo });
        mnuHelp.Name = "mnuHelp";
        mnuHelp.Text = "❓ Help";
        // 
        // mnuHelpHelp
        // 
        mnuHelpHelp.Name = "mnuHelpHelp";
        mnuHelpHelp.Text = "Help";
        mnuHelpHelp.Click += mnuHelpHelp_Click;
        // 
        // mnuHelpInfo
        // 
        mnuHelpInfo.Name = "mnuHelpInfo";
        mnuHelpInfo.Text = "Info";
        mnuHelpInfo.Click += mnuHelpInfo_Click;
        // 
        // lblQuery
        // 
        lblQuery.AutoSize = true;
        lblQuery.Location = new Point(8, 36);
        lblQuery.Name = "lblQuery";
        lblQuery.Text = "Zoekterm:";
        // 
        // cmbQuery
        // 
        cmbQuery.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        cmbQuery.DropDownStyle = ComboBoxStyle.DropDown;
        cmbQuery.Location = new Point(75, 32);
        cmbQuery.Name = "cmbQuery";
        cmbQuery.Size = new Size(829, 23);
        cmbQuery.TabIndex = 1;
        cmbQuery.KeyDown += cmbQuery_KeyDown;
        // 
        // btnSearch
        // 
        btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnSearch.Location = new Point(912, 31);
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
        btnCancel.Location = new Point(1006, 31);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(86, 27);
        btnCancel.TabIndex = 3;
        btnCancel.Text = "Annuleren";
        btnCancel.UseVisualStyleBackColor = true;
        btnCancel.Click += btnCancel_Click;
        // 
        // chkShowPreview
        // 
        chkShowPreview.AutoSize = true;
        chkShowPreview.Location = new Point(8, 67);
        chkShowPreview.Name = "chkShowPreview";
        chkShowPreview.Text = "Toon voorbeeldvenster";
        chkShowPreview.UseVisualStyleBackColor = true;
        // 
        // chkUseDateRange
        // 
        chkUseDateRange.AutoSize = true;
        chkUseDateRange.Checked = true;
        chkUseDateRange.CheckState = CheckState.Checked;
        chkUseDateRange.Location = new Point(168, 67);
        chkUseDateRange.Name = "chkUseDateRange";
        chkUseDateRange.Text = "Gebruik datum";
        chkUseDateRange.UseVisualStyleBackColor = true;
        // 
        // dtpFrom
        // 
        dtpFrom.Format = DateTimePickerFormat.Short;
        dtpFrom.Location = new Point(278, 64);
        dtpFrom.Name = "dtpFrom";
        dtpFrom.Size = new Size(94, 23);
        dtpFrom.TabIndex = 5;
        // 
        // dtpTo
        // 
        dtpTo.Format = DateTimePickerFormat.Short;
        dtpTo.Location = new Point(378, 64);
        dtpTo.Name = "dtpTo";
        dtpTo.Size = new Size(94, 23);
        dtpTo.TabIndex = 6;
        // 
        // pnlFilters
        // 
        pnlFilters.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        pnlFilters.Location = new Point(8, 96);
        pnlFilters.Name = "pnlFilters";
        pnlFilters.Size = new Size(1084, 27);
        pnlFilters.TabIndex = 7;
        pnlFilters.Controls.Add(txtColMailbox);
        pnlFilters.Controls.Add(txtColFolderPath);
        pnlFilters.Controls.Add(txtColDate);
        pnlFilters.Controls.Add(txtColSubject);
        pnlFilters.Controls.Add(txtColSender);
        pnlFilters.Controls.Add(txtColRecipients);
        // 
        // txtColMailbox
        // 
        txtColMailbox.Name = "txtColMailbox";
        txtColMailbox.PlaceholderText = "Filter: Mailbox";
        txtColMailbox.TabIndex = 0;
        // 
        // txtColFolderPath
        // 
        txtColFolderPath.Name = "txtColFolderPath";
        txtColFolderPath.PlaceholderText = "Filter: Map";
        txtColFolderPath.TabIndex = 1;
        // 
        // txtColDate
        // 
        txtColDate.Name = "txtColDate";
        txtColDate.PlaceholderText = "Filter: Datum";
        txtColDate.TabIndex = 2;
        // 
        // txtColSubject
        // 
        txtColSubject.Name = "txtColSubject";
        txtColSubject.PlaceholderText = "Filter: Onderwerp";
        txtColSubject.TabIndex = 3;
        // 
        // txtColSender
        // 
        txtColSender.Name = "txtColSender";
        txtColSender.PlaceholderText = "Filter: Afzender";
        txtColSender.TabIndex = 4;
        // 
        // txtColRecipients
        // 
        txtColRecipients.Name = "txtColRecipients";
        txtColRecipients.PlaceholderText = "Filter: Geadresseerde";
        txtColRecipients.TabIndex = 5;
        // 
        // splitContainer
        // 
        splitContainer.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        splitContainer.Location = new Point(8, 127);
        splitContainer.Name = "splitContainer";
        splitContainer.Orientation = Orientation.Vertical;
        splitContainer.Panel1MinSize = 200;
        splitContainer.Panel2MinSize = 150;
        splitContainer.Size = new Size(1084, 565);
        splitContainer.SplitterDistance = 580;
        splitContainer.SplitterWidth = 6;
        splitContainer.TabIndex = 8;
        splitContainer.Panel2Collapsed = true;
        splitContainer.Panel1.Controls.Add(dgvResults);
        splitContainer.Panel2.Controls.Add(pnlPreview);
        // 
        // dgvResults
        // 
        dgvResults.AllowUserToAddRows = false;
        dgvResults.AllowUserToDeleteRows = false;
        dgvResults.AllowUserToOrderColumns = true;
        dgvResults.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResults.Dock = DockStyle.Fill;
        dgvResults.Name = "dgvResults";
        dgvResults.ReadOnly = true;
        dgvResults.RowHeadersVisible = false;
        dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvResults.TabIndex = 0;
        dgvResults.CellDoubleClick += dgvResults_CellDoubleClick;
        // 
        // pnlPreviewTop
        // 
        pnlPreviewTop.Controls.Add(btnClosePreview);
        pnlPreviewTop.Controls.Add(btnCopyPreview);
        pnlPreviewTop.Dock = DockStyle.Top;
        pnlPreviewTop.Height = 34;
        pnlPreviewTop.Name = "pnlPreviewTop";
        pnlPreviewTop.Size = new Size(498, 34);
        pnlPreviewTop.TabIndex = 0;
        // 
        // pnlPreview
        // 
        pnlPreview.BorderStyle = BorderStyle.FixedSingle;
        pnlPreview.Controls.Add(txtPreview);
        pnlPreview.Controls.Add(pnlPreviewTop);
        pnlPreview.Dock = DockStyle.Fill;
        pnlPreview.Name = "pnlPreview";
        pnlPreview.TabIndex = 0;
        // 
        // btnClosePreview
        // 
        btnClosePreview.Anchor = AnchorStyles.Top | AnchorStyles.Left;
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
        btnCopyPreview.Location = new Point(377, 5);
        btnCopyPreview.Name = "btnCopyPreview";
        btnCopyPreview.Size = new Size(115, 24);
        btnCopyPreview.TabIndex = 1;
        btnCopyPreview.Text = "Kopieer voorbeeld";
        btnCopyPreview.UseVisualStyleBackColor = true;
        btnCopyPreview.Click += btnCopyPreview_Click;
        // 
        // txtPreview
        // 
        txtPreview.Dock = DockStyle.Fill;
        txtPreview.Multiline = true;
        txtPreview.Name = "txtPreview";
        txtPreview.ReadOnly = true;
        txtPreview.ScrollBars = ScrollBars.Both;
        txtPreview.TabIndex = 1;
        // 
        // statusStrip1
        // 
        statusStrip1.Items.AddRange(new ToolStripItem[] { toolStripProgressBar, toolStripStatusLabel1 });
        statusStrip1.Dock = DockStyle.Bottom;
        statusStrip1.Name = "statusStrip1";
        statusStrip1.Size = new Size(1100, 22);
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
        toolStripStatusLabel1.Spring = true;
        toolStripStatusLabel1.TextAlign = ContentAlignment.MiddleLeft;
        toolStripStatusLabel1.Text = "Klaar voor zoekopdracht";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1100, 718);
        MainMenuStrip = menuStrip1;
        Controls.Add(splitContainer);
        Controls.Add(pnlFilters);
        Controls.Add(dtpTo);
        Controls.Add(dtpFrom);
        Controls.Add(chkUseDateRange);
        Controls.Add(chkShowPreview);
        Controls.Add(btnCancel);
        Controls.Add(btnSearch);
        Controls.Add(cmbQuery);
        Controls.Add(lblQuery);
        Controls.Add(menuStrip1);
        Controls.Add(statusStrip1);
        MinimumSize = new Size(800, 500);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Outlook Classic Search";
        Load += Form1_Load;
        FormClosing += Form1_FormClosing;
        menuStrip1.ResumeLayout(false);
        menuStrip1.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer).EndInit();
        splitContainer.Panel1.ResumeLayout(false);
        splitContainer.Panel2.ResumeLayout(false);
        splitContainer.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResults).EndInit();
        pnlPreviewTop.ResumeLayout(false);
        pnlPreview.ResumeLayout(false);
        pnlPreview.PerformLayout();
        statusStrip1.ResumeLayout(false);
        statusStrip1.PerformLayout();
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
    private CheckBox chkShowPreview;
    private CheckBox chkUseDateRange;
    private DateTimePicker dtpFrom;
    private DateTimePicker dtpTo;
    private Panel pnlFilters;
    private TextBox txtColMailbox;
    private TextBox txtColFolderPath;
    private TextBox txtColDate;
    private TextBox txtColSubject;
    private TextBox txtColSender;
    private TextBox txtColRecipients;
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
}
