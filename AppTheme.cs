using System.Runtime.InteropServices;

namespace OutlookClassicSearch;

/// <summary>
/// Centrale theming-klasse. Roep Apply(form) aan na InitializeComponent()
/// en gebruik ApplyPrimaryStyle(button) voor de primaire actieknop.
/// </summary>
internal static class AppTheme
{
    // ── Kleurenpalet (licht, plat) ────────────────────────────────────────────
    public static readonly Color Background     = Color.FromArgb(245, 246, 247);
    public static readonly Color Surface        = Color.White;
    public static readonly Color Accent         = Color.FromArgb(0, 120, 212);
    public static readonly Color TextPrimary    = Color.FromArgb(32,  32,  32);
    public static readonly Color TextMuted      = Color.FromArgb(100, 100, 100);
    public static readonly Color Border         = Color.FromArgb(213, 213, 213);
    public static readonly Color StatusBar      = Color.FromArgb(228, 228, 228);
    public static readonly Color GridHeader     = Color.FromArgb(235, 236, 237);
    public static readonly Color GridAltRow     = Color.FromArgb(249, 250, 252);
    public static readonly Color GridSelect     = Color.FromArgb(0,   120, 212);
    public static readonly Color GridSelectText = Color.White;

    // ── Lettertypen ──────────────────────────────────────────────────────────
    public static readonly Font DefaultFont = new("Segoe UI", 9f);
    public static readonly Font BoldFont    = new("Segoe UI", 9f, FontStyle.Bold);

    // ── Publieke API ─────────────────────────────────────────────────────────

    /// <summary>Past het thema toe op het formulier en alle onderliggende controls.</summary>
    public static void Apply(Form form)
    {
        form.BackColor = Background;
        form.Font      = DefaultFont;
        ApplyDwmLightFrame(form);
        StyleAllControls(form);
    }

    /// <summary>Stijlt een knop als primaire actieknop (accentkleur).</summary>
    public static void ApplyPrimaryStyle(Button button)
    {
        button.UseVisualStyleBackColor = false;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = Accent;
        button.ForeColor = Color.White;
        button.Font      = BoldFont;
    }

    // ── Recursieve walker ─────────────────────────────────────────────────────

    private static void StyleAllControls(Control parent)
    {
        foreach (Control ctrl in parent.Controls)
        {
            StyleControl(ctrl);

            // ToolStrip-items zitten niet in de gewone Controls-collectie
            if (ctrl is ToolStrip ts)
                StyleToolStripItems(ts.Items);

            if (ctrl.HasChildren)
                StyleAllControls(ctrl);
        }
    }

    private static void StyleControl(Control ctrl)
    {
        switch (ctrl)
        {
            case MenuStrip ms:
                StyleMenuStrip(ms);
                break;

            case StatusStrip ss:
                StyleStatusStrip(ss);
                break;

            case DataGridView dgv:
                StyleDataGridView(dgv);
                break;

            case Button btn:
                StyleButton(btn);
                break;

            case TextBox tb:
                StyleTextBox(tb);
                break;

            case ComboBox cmb:
                StyleComboBox(cmb);
                break;

            case CheckBox chk:
                chk.ForeColor = TextPrimary;
                break;

            case Label lbl:
                lbl.ForeColor = TextPrimary;
                break;

            case Panel pnl:
                pnl.BackColor = Background;
                break;

            case SplitContainer sc:
                sc.BackColor = Border;  // splitter-balk krijgt subtiele scheiding
                break;

            case NumericUpDown nud:
                nud.BackColor = Surface;
                nud.ForeColor = TextPrimary;
                break;

            case CheckedListBox clb:
                clb.BackColor = Surface;
                clb.ForeColor = TextPrimary;
                break;

            case ListBox lb:
                lb.BackColor = Surface;
                lb.ForeColor = TextPrimary;
                break;

            case TreeView tv:
                tv.BackColor = Surface;
                tv.ForeColor = TextPrimary;
                break;
        }
    }

    // ── Per-control stijlen ───────────────────────────────────────────────────

    private static void StyleButton(Button btn)
    {
        btn.UseVisualStyleBackColor = false;
        btn.FlatStyle = FlatStyle.Flat;
        btn.FlatAppearance.BorderColor = Border;
        btn.FlatAppearance.BorderSize  = 1;
        btn.BackColor = Surface;
        btn.ForeColor = TextPrimary;
    }

    private static void StyleTextBox(TextBox tb)
    {
        tb.BackColor = Surface;
        tb.ForeColor = TextPrimary;
        // Vervang de klassieke 3D-rand door een plattere variant;
        // laat BorderStyle.None met rust (bewust zo ingesteld).
        if (tb.BorderStyle == BorderStyle.Fixed3D)
            tb.BorderStyle = BorderStyle.FixedSingle;
    }

    private static void StyleComboBox(ComboBox cmb)
    {
        cmb.BackColor  = Surface;
        cmb.ForeColor  = TextPrimary;
        cmb.FlatStyle  = FlatStyle.Flat;
    }

    private static void StyleMenuStrip(MenuStrip ms)
    {
        ms.BackColor = Background;
        ms.ForeColor = TextPrimary;
        ms.Renderer  = new ToolStripProfessionalRenderer(new ModernColorTable());
    }

    private static void StyleStatusStrip(StatusStrip ss)
    {
        ss.BackColor  = StatusBar;
        ss.ForeColor  = TextMuted;
        ss.Renderer   = new ToolStripProfessionalRenderer(new ModernColorTable());
        ss.SizingGrip = false;
    }

    private static void StyleDataGridView(DataGridView dgv)
    {
        dgv.EnableHeadersVisualStyles = false;
        dgv.BackgroundColor           = Background;
        dgv.BorderStyle               = BorderStyle.None;
        dgv.GridColor                 = Border;
        dgv.RowHeadersVisible         = false;

        dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor         = GridHeader,
            ForeColor         = TextPrimary,
            Font              = BoldFont,
            SelectionBackColor = GridHeader,
            SelectionForeColor = TextPrimary,
            Padding           = new Padding(4, 0, 4, 0)
        };

        dgv.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor          = Surface,
            ForeColor          = TextPrimary,
            SelectionBackColor = GridSelect,
            SelectionForeColor = GridSelectText
        };

        dgv.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor          = GridAltRow,
            ForeColor          = TextPrimary,
            SelectionBackColor = GridSelect,
            SelectionForeColor = GridSelectText
        };
    }

    private static void StyleToolStripItems(ToolStripItemCollection items)
    {
        foreach (ToolStripItem item in items)
        {
            item.BackColor = Background;
            item.ForeColor = TextPrimary;

            if (item is ToolStripDropDownItem ddi)
                StyleToolStripItems(ddi.DropDownItems);
        }
    }

    // ── Windows 11 DWM – lichte titelbalktint ────────────────────────────────

    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    [DllImport("dwmapi.dll", PreserveSig = false)]
    private static extern void DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    private static void ApplyDwmLightFrame(Form form)
    {
        void Apply()
        {
            try
            {
                int value = 0; // 0 = licht, 1 = donker
                DwmSetWindowAttribute(form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
            }
            catch
            {
                // Niet ondersteund op oudere Windows-versies – negeer
            }
        }

        if (form.IsHandleCreated)
            Apply();
        else
            form.HandleCreated += (_, _) => Apply();
    }

    // ── Kleurentabel voor menu- en statusbalk ─────────────────────────────────

    private sealed class ModernColorTable : ProfessionalColorTable
    {
        public override Color MenuStripGradientBegin              => Background;
        public override Color MenuStripGradientEnd                => Background;
        public override Color MenuBorder                          => Border;
        public override Color MenuItemBorder                      => Accent;
        public override Color MenuItemSelected                    => Color.FromArgb(229, 243, 255);
        public override Color MenuItemSelectedGradientBegin       => Color.FromArgb(229, 243, 255);
        public override Color MenuItemSelectedGradientEnd         => Color.FromArgb(229, 243, 255);
        public override Color MenuItemPressedGradientBegin        => Color.FromArgb(204, 232, 255);
        public override Color MenuItemPressedGradientEnd          => Color.FromArgb(204, 232, 255);
        public override Color ToolStripDropDownBackground         => Surface;
        public override Color ImageMarginGradientBegin            => Surface;
        public override Color ImageMarginGradientMiddle           => Surface;
        public override Color ImageMarginGradientEnd              => Surface;
        public override Color StatusStripGradientBegin            => StatusBar;
        public override Color StatusStripGradientEnd              => StatusBar;
        public override Color SeparatorDark                       => Border;
        public override Color SeparatorLight                      => Surface;
    }
}
