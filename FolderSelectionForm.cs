namespace OutlookClassicSearch;

internal sealed class FolderSelectionForm : Form
{
    private readonly TreeView _tree = new();
    private readonly Button _btnOk = new();
    private readonly Button _btnCancel = new();
    private readonly Label _lblInfo = new();
    private readonly HashSet<string> _preselected;
    private readonly Func<CancellationToken, Task<IReadOnlyList<MailboxFolderRoot>>>? _rootsLoader;
    private readonly CancellationTokenSource _loadCts = new();
    private bool _loading;

    public FolderSelectionForm(
        string title,
        IReadOnlyList<MailboxFolderRoot> roots,
        IEnumerable<string> preselectedEntryIds,
        Func<CancellationToken, Task<IReadOnlyList<MailboxFolderRoot>>>? rootsLoader = null)
    {
        _preselected = new HashSet<string>(preselectedEntryIds, StringComparer.OrdinalIgnoreCase);
        _rootsLoader = rootsLoader;

        Text = title;
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 620);
        MinimumSize = new Size(600, 450);

        _lblInfo.Dock = DockStyle.Top;
        _lblInfo.Height = 32;
        _lblInfo.Text = "Vink mappen aan en klik op Opslaan.";
        _lblInfo.Padding = new Padding(8, 8, 8, 0);

        _tree.Dock = DockStyle.Fill;
        _tree.CheckBoxes = true;
        _tree.AfterCheck += TreeAfterCheck;

        var contentPanel = new Panel { Dock = DockStyle.Fill };
        contentPanel.Controls.Add(_tree);

        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 52 };
        var buttonsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            Width = 230,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(0, 10, 8, 0),
            WrapContents = false
        };

        _btnOk.Text = "Opslaan";
        _btnOk.Width = 100;
        _btnOk.Height = 30;
        _btnOk.Margin = new Padding(8, 0, 0, 0);
        _btnOk.Click += (_, _) => DialogResult = DialogResult.OK;

        _btnCancel.Text = "Annuleren";
        _btnCancel.Width = 100;
        _btnCancel.Height = 30;
        _btnCancel.Margin = new Padding(8, 0, 0, 0);
        _btnCancel.Click += (_, _) => DialogResult = DialogResult.Cancel;

        buttonsPanel.Controls.Add(_btnCancel);
        buttonsPanel.Controls.Add(_btnOk);
        bottomPanel.Controls.Add(buttonsPanel);

        Controls.Add(contentPanel);
        Controls.Add(bottomPanel);
        Controls.Add(_lblInfo);

        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        LoadTree(roots);

        if (roots.Count == 0 && _rootsLoader is not null)
        {
            Shown += async (_, _) => await LoadRootsAsync();
        }

        FormClosing += (_, _) => _loadCts.Cancel();
    }

    public IReadOnlyList<string> GetSelectedEntryIds()
    {
        var selected = new List<string>();
        foreach (TreeNode root in _tree.Nodes)
        {
            CollectChecked(root, selected);
        }

        return selected;
    }

    private void LoadTree(IReadOnlyList<MailboxFolderRoot> roots)
    {
        _tree.Nodes.Clear();

        foreach (var root in roots)
        {
            var rootNode = new TreeNode(root.DisplayName)
            {
                Tag = root,
                Checked = _preselected.Contains(root.EntryId)
            };

            AddChildren(rootNode, root.Children);
            _tree.Nodes.Add(rootNode);
            rootNode.Expand();
        }
    }

    private async Task LoadRootsAsync()
    {
        if (_rootsLoader is null || _loading)
        {
            return;
        }

        _loading = true;
        _btnOk.Enabled = false;
        _tree.Enabled = false;
        _lblInfo.Text = "Mappenstructuur laden...";

        try
        {
            var loadedRoots = await _rootsLoader(_loadCts.Token);
            LoadTree(loadedRoots);
            _lblInfo.Text = "Vink mappen aan en klik op Opslaan.";
        }
        catch (OperationCanceledException)
        {
            _lblInfo.Text = "Laden geannuleerd.";
        }
        catch
        {
            _lblInfo.Text = "Laden mislukt. Sluit dit venster en probeer opnieuw.";
        }
        finally
        {
            _btnOk.Enabled = true;
            _tree.Enabled = true;
            _loading = false;
        }
    }

    private void AddChildren(TreeNode parent, IReadOnlyList<MailboxFolderNode> children)
    {
        foreach (var child in children)
        {
            var node = new TreeNode(child.DisplayName)
            {
                Tag = child,
                Checked = _preselected.Contains(child.EntryId)
            };

            AddChildren(node, child.Children);
            parent.Nodes.Add(node);
        }
    }

    private void TreeAfterCheck(object? sender, TreeViewEventArgs e)
    {
        if (e.Action == TreeViewAction.Unknown || e.Node is null)
        {
            return;
        }

        foreach (TreeNode child in e.Node.Nodes)
        {
            child.Checked = e.Node.Checked;
        }
    }

    private static void CollectChecked(TreeNode node, List<string> selected)
    {
        if (node.Checked)
        {
            switch (node.Tag)
            {
                case MailboxFolderRoot root:
                    selected.Add(root.EntryId);
                    break;
                case MailboxFolderNode child:
                    selected.Add(child.EntryId);
                    break;
            }
        }

        foreach (TreeNode child in node.Nodes)
        {
            CollectChecked(child, selected);
        }
    }
}
