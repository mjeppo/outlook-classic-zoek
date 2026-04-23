using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OutlookClassicSearch;

/// <summary>
/// Centrale theming-klasse. Roep Apply(form) aan na InitializeComponent()
/// en gebruik ApplyPrimaryStyle(button) voor de primaire actieknop.
/// </summary>
internal static class AppTheme
{
    // ── Theme modus ──────────────────────────────────────────────────────────
    public enum ThemeMode
    {
        Light,
        Dark,
        Auto
    }

    private static ThemeMode _currentMode = ThemeMode.Light;
    private static bool _isDarkTheme = false;

    /// <summary>Huidige theme modus (Light, Dark, of Auto)</summary>
    public static ThemeMode CurrentMode
    {
        get => _currentMode;
        set
        {
            _currentMode = value;
            UpdateEffectiveTheme();
        }
    }

    /// <summary>Is het huidige effectieve thema donker?</summary>
    public static bool IsDarkTheme => _isDarkTheme;

    /// <summary>Event dat wordt aangeroepen wanneer het thema verandert</summary>
    public static event EventHandler? ThemeChanged;

    /// <summary>Update het effectieve thema op basis van de modus</summary>
    private static void UpdateEffectiveTheme()
    {
        bool wasDark = _isDarkTheme;
        _isDarkTheme = _currentMode switch
        {
            ThemeMode.Dark => true,
            ThemeMode.Auto => IsSystemDarkMode(),
            _ => false
        };

        // Als het thema daadwerkelijk is veranderd, roep dan het event aan
        if (wasDark != _isDarkTheme)
        {
            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }
    }

    /// <summary>Wijzig het thema en update MaterialSkin + alle forms</summary>
    public static void ChangeTheme(ThemeMode newMode)
    {
        CurrentMode = newMode;

        // Update MaterialSkinManager
        var materialSkinManager = MaterialSkin.MaterialSkinManager.Instance;
        materialSkinManager.Theme = _isDarkTheme 
            ? MaterialSkin.MaterialSkinManager.Themes.DARK 
            : MaterialSkin.MaterialSkinManager.Themes.LIGHT;

        // Trigger het ThemeChanged event om forms te laten refreshen
        ThemeChanged?.Invoke(null, EventArgs.Empty);
    }

    /// <summary>Detecteert of Windows in donkere modus staat</summary>
    private static bool IsSystemDarkMode()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            return value is int intValue && intValue == 0;
        }
        catch
        {
            return false;
        }
    }

    // ── Kleurenpalet (dynamisch op basis van thema) ──────────────────────────

    // Licht thema kleuren
    private static class LightColors
    {
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
    }

    // Donker thema kleuren
    private static class DarkColors
    {
        public static readonly Color Background     = Color.FromArgb(32, 32, 32);
        public static readonly Color Surface        = Color.FromArgb(45, 45, 48);
        public static readonly Color Accent         = Color.FromArgb(0, 120, 212);
        public static readonly Color TextPrimary    = Color.FromArgb(240, 240, 240);
        public static readonly Color TextMuted      = Color.FromArgb(160, 160, 160);
        public static readonly Color Border         = Color.FromArgb(63, 63, 70);
        public static readonly Color StatusBar      = Color.FromArgb(37, 37, 38);
        public static readonly Color GridHeader     = Color.FromArgb(51, 51, 55);
        public static readonly Color GridAltRow     = Color.FromArgb(40, 40, 43);
        public static readonly Color GridSelect     = Color.FromArgb(0, 120, 212);
        public static readonly Color GridSelectText = Color.White;
    }

    // Publieke properties die het juiste thema retourneren
    public static Color Background     => _isDarkTheme ? DarkColors.Background     : LightColors.Background;
    public static Color Surface        => _isDarkTheme ? DarkColors.Surface        : LightColors.Surface;
    public static Color Accent         => _isDarkTheme ? DarkColors.Accent         : LightColors.Accent;
    public static Color TextPrimary    => _isDarkTheme ? DarkColors.TextPrimary    : LightColors.TextPrimary;
    public static Color TextMuted      => _isDarkTheme ? DarkColors.TextMuted      : LightColors.TextMuted;
    public static Color Border         => _isDarkTheme ? DarkColors.Border         : LightColors.Border;
    public static Color StatusBar      => _isDarkTheme ? DarkColors.StatusBar      : LightColors.StatusBar;
    public static Color GridHeader     => _isDarkTheme ? DarkColors.GridHeader     : LightColors.GridHeader;
    public static Color GridAltRow     => _isDarkTheme ? DarkColors.GridAltRow     : LightColors.GridAltRow;
    public static Color GridSelect     => _isDarkTheme ? DarkColors.GridSelect     : LightColors.GridSelect;
    public static Color GridSelectText => _isDarkTheme ? DarkColors.GridSelectText : LightColors.GridSelectText;

    // ── Lettertypen ──────────────────────────────────────────────────────────
    public static readonly Font DefaultFont = new("Segoe UI", 9f);
    public static readonly Font BoldFont    = new("Segoe UI", 9f, FontStyle.Bold);
    public static readonly Font ButtonFont  = new("Segoe UI", 9f, FontStyle.Regular);

    // ── Publieke API ─────────────────────────────────────────────────────────

    /// <summary>Past het thema toe op het formulier en alle onderliggende controls.</summary>
    public static void Apply(Form form)
    {
        form.BackColor = Background;
        form.Font      = DefaultFont;
        ApplyDwmLightFrame(form);
        StyleAllControls(form);
    }

    /// <summary>Stijlt een knop als primaire actieknop (accentkleur met verhoogd effect).</summary>
    public static void ApplyPrimaryStyle(Button button)
    {
        // Verwijder eventuele bestaande hover handlers
        RemoveHoverHandlers(button);

        button.UseVisualStyleBackColor = false;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = Accent;
        button.ForeColor = Color.White;
        button.Font      = ButtonFont;
        button.Cursor    = Cursors.Hand;
        // Geen padding - behoud de Designer size

        // Hover effect
        EventHandler onEnter = (s, e) => { button.BackColor = Color.FromArgb(0, 100, 190); };
        EventHandler onLeave = (s, e) => { button.BackColor = Accent; };

        button.MouseEnter += onEnter;
        button.MouseLeave += onLeave;

        // Sla handlers op voor later verwijderen
        button.Tag = new Tuple<EventHandler, EventHandler>(onEnter, onLeave);
    }

    /// <summary>Stijlt een knop als secundaire actieknop (outline stijl).</summary>
    public static void ApplySecondaryStyle(Button button)
    {
        // Verwijder eventuele bestaande hover handlers
        RemoveHoverHandlers(button);

        button.UseVisualStyleBackColor = false;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = Accent;
        button.BackColor = Surface;
        button.ForeColor = Accent;
        button.Font      = ButtonFont;
        button.Cursor    = Cursors.Hand;
        // Geen padding - behoud de Designer size

        // Hover effect (dynamisch op basis van thema)
        Color hoverColor = _isDarkTheme 
            ? Color.FromArgb(55, 55, 58)   // Donker hover
            : Color.FromArgb(240, 245, 250); // Licht hover

        EventHandler onEnter = (s, e) => { button.BackColor = hoverColor; };
        EventHandler onLeave = (s, e) => { button.BackColor = Surface; };

        button.MouseEnter += onEnter;
        button.MouseLeave += onLeave;

        // Sla handlers op voor later verwijderen
        button.Tag = new Tuple<EventHandler, EventHandler>(onEnter, onLeave);
    }

    /// <summary>Verwijdert bestaande hover handlers van een button</summary>
    private static void RemoveHoverHandlers(Button button)
    {
        if (button.Tag is Tuple<EventHandler, EventHandler> handlers)
        {
            button.MouseEnter -= handlers.Item1;
            button.MouseLeave -= handlers.Item2;
            button.Tag = null;
        }
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

            case MaterialSkin.Controls.MaterialSwitch sw:
                StyleMaterialSwitch(sw);
                break;

            case CheckBox chk:
                StyleCheckBox(chk);
                break;

            case Label lbl:
                lbl.ForeColor = TextPrimary;
                lbl.BackColor = Color.Transparent;
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
        // Moderne flat button met subtiele rand
        btn.UseVisualStyleBackColor = false;
        btn.FlatStyle = FlatStyle.Flat;
        btn.FlatAppearance.BorderColor = Border;
        btn.FlatAppearance.BorderSize  = 1;
        btn.BackColor = Surface;
        btn.ForeColor = TextPrimary;
        btn.Font      = ButtonFont;
        btn.Cursor    = Cursors.Hand;
        // Geen padding hier - laat de Designer sizes intact

        // Hover effect voor betere UX
        btn.MouseEnter += (s, e) => 
        { 
            if (btn.BackColor == Surface)
                btn.BackColor = Color.FromArgb(245, 245, 245);
        };
        btn.MouseLeave += (s, e) => 
        { 
            if (btn.BackColor == Color.FromArgb(245, 245, 245))
                btn.BackColor = Surface;
        };
    }

    private static void StyleTextBox(TextBox tb)
    {
        tb.BackColor = Surface;
        tb.ForeColor = TextPrimary;
        tb.Font      = DefaultFont;

        // Gebruik FixedSingle voor een rand - accepteer de zwarte kleur voorlopig
        if (tb.BorderStyle == BorderStyle.Fixed3D)
            tb.BorderStyle = BorderStyle.FixedSingle;

        // De wrapper aanpak werkt niet goed, dus laten we het simpel houden
        // De zwarte rand is niet ideaal maar werkt wel betrouwbaar
    }

    private static void StyleComboBox(ComboBox cmb)
    {
        cmb.BackColor  = Surface;
        cmb.ForeColor  = TextPrimary;
        cmb.Font       = DefaultFont;
        cmb.FlatStyle  = FlatStyle.Flat;
    }

    public static void StyleMaterialSwitch(MaterialSkin.Controls.MaterialSwitch sw)
    {
        // MaterialSwitch heeft zijn eigen styling via MaterialSkin
        // We gebruiken labels voor de tekst, dus alleen de Depth instellen
        sw.Depth = 0; // 0 = primary depth voor betere visuele hiërarchie

        // MaterialSwitch rendering fix voor dark mode
        // Forceer de switch om de parent background color te gebruiken
        sw.BackColor = Background;

        // Force invalidate om de rendering te triggeren
        sw.Invalidate();
    }

    private static void StyleCheckBox(CheckBox chk)
    {
        chk.ForeColor = TextPrimary;
        chk.BackColor = Color.Transparent;
        chk.Font      = DefaultFont;
        chk.FlatStyle = FlatStyle.Flat;
        chk.Cursor    = Cursors.Hand;
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

        // Hover kleuren - afhankelijk van thema
        public override Color MenuItemSelected                    => _isDarkTheme 
            ? Color.FromArgb(60, 60, 60)   // Donkergrijs hover in dark mode
            : Color.FromArgb(229, 243, 255); // Lichtblauw hover in light mode

        public override Color MenuItemSelectedGradientBegin       => MenuItemSelected;
        public override Color MenuItemSelectedGradientEnd         => MenuItemSelected;

        // Pressed kleuren
        public override Color MenuItemPressedGradientBegin        => _isDarkTheme
            ? Color.FromArgb(50, 50, 50)   // Nog donkerder in dark mode
            : Color.FromArgb(204, 232, 255); // Donkerder blauw in light mode

        public override Color MenuItemPressedGradientEnd          => MenuItemPressedGradientBegin;

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
