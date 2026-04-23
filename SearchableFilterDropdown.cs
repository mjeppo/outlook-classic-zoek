using System.Windows.Forms;

namespace OutlookClassicSearch;

/// <summary>
/// Zoekbare dropdown filter met live filtering terwijl je typt.
/// </summary>
internal sealed class SearchableFilterDropdown : Form
{
    private readonly TextBox _searchBox;
    private readonly CheckedListBox _listBox;
    private readonly List<string> _allItems;
    private readonly HashSet<string> _selectedItems;

    public SearchableFilterDropdown(List<string> items, HashSet<string> selectedItems, Point location)
    {
        _allItems = items;
        _selectedItems = selectedItems;

        FormBorderStyle = FormBorderStyle.FixedSingle;
        StartPosition = FormStartPosition.Manual;
        Location = location;
        Width = 300;
        Height = 350;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = AppTheme.Surface;
        ControlBox = true;
        MaximizeBox = false;
        MinimizeBox = false;
        Text = Strings.IsEnglish ? "Select folders" : "Selecteer mappen";

        _searchBox = new TextBox
        {
            Dock = DockStyle.Top,
            PlaceholderText = Strings.IsEnglish ? "Type to search..." : "Type om te zoeken...",
            BackColor = AppTheme.Background,
            ForeColor = AppTheme.TextPrimary,
            BorderStyle = BorderStyle.FixedSingle,
        };
        _searchBox.TextChanged += SearchBox_TextChanged;
        _searchBox.KeyDown += SearchBox_KeyDown;

        _listBox = new CheckedListBox
        {
            Dock = DockStyle.Fill,
            CheckOnClick = true,
            IntegralHeight = false,
            BackColor = AppTheme.Surface,
            ForeColor = AppTheme.TextPrimary,
            BorderStyle = BorderStyle.None,
            HorizontalScrollbar = true,
        };
        _listBox.ItemCheck += ListBox_ItemCheck;
        _listBox.KeyDown += (s, e) => 
        {
            if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        };

        Controls.Add(_listBox);
        Controls.Add(_searchBox);

        PopulateList(string.Empty);
        _searchBox.Focus();
    }

    private void SearchBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
        else if (e.KeyCode == Keys.Down && _listBox.Items.Count > 0)
        {
            _listBox.Focus();
            _listBox.SelectedIndex = 0;
        }
    }

    private void SearchBox_TextChanged(object? sender, EventArgs e)
    {
        PopulateList(_searchBox.Text);
    }

    private void PopulateList(string filter)
    {
        _listBox.ItemCheck -= ListBox_ItemCheck;
        _listBox.Items.Clear();

        var filteredItems = string.IsNullOrWhiteSpace(filter)
            ? _allItems
            : _allItems.Where(item => item.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

        foreach (var item in filteredItems)
        {
            int index = _listBox.Items.Add(item);
            if (_selectedItems.Contains(item))
            {
                _listBox.SetItemChecked(index, true);
            }
        }

        _listBox.ItemCheck += ListBox_ItemCheck;
    }

    private void ListBox_ItemCheck(object? sender, ItemCheckEventArgs e)
    {
        if (_listBox.Items[e.Index] is string item)
        {
            // ItemCheck fires BEFORE the check state changes, so schedule the update
            BeginInvoke(() =>
            {
                if (e.NewValue == CheckState.Checked)
                {
                    _selectedItems.Add(item);
                }
                else
                {
                    _selectedItems.Remove(item);
                }
                OnSelectionChanged();
            });
        }
    }

    public event EventHandler? SelectionChanged;

    private void OnSelectionChanged()
    {
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    protected override void OnDeactivate(EventArgs e)
    {
        base.OnDeactivate(e);
        DialogResult = DialogResult.OK;
        Close();
    }
}
