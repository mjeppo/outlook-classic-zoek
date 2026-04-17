using System.Runtime.InteropServices;

namespace OutlookClassicSearch;

/// <summary>
/// TextBox met een custom border kleur die past bij het AppTheme en verticaal gecentreerde tekst
/// </summary>
internal class BorderedTextBox : TextBox
{
    private const int WM_PAINT = 0xF;
    private const int WM_NCPAINT = 0x85;

    public BorderedTextBox()
    {
        BorderStyle = BorderStyle.FixedSingle;
        // Multiline om hoogte van 27px te kunnen forceren
        Multiline = true;
        // Voorkom Enter voor nieuwe regels
        AcceptsReturn = false;
        // Gebruik iets grotere font voor betere centrering
        Font = new Font("Segoe UI", 9.75f);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        // Bereken de benodigde top margin voor verticale centrering
        CenterTextVertically();
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        CenterTextVertically();
    }

    private void CenterTextVertically()
    {
        if (!IsHandleCreated) return;

        // Bereken de teksthoogte
        using var g = CreateGraphics();
        var textSize = g.MeasureString("Ag", Font);
        var textHeight = (int)Math.Ceiling(textSize.Height);

        // Bereken de top margin om tekst te centreren in 27px hoogte
        var topMargin = Math.Max(0, (27 - textHeight) / 2 - 2); // -2 voor border compensatie

        // Pas padding aan
        Padding = new Padding(2, topMargin, 2, 0);
    }

    protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
    {
        // Forceer altijd hoogte van 27px
        base.SetBoundsCore(x, y, width, 27, specified);
    }

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);

        // Teken een gekleurde rand na de standaard paint
        if (m.Msg == WM_NCPAINT || m.Msg == WM_PAINT)
        {
            using var g = Graphics.FromHwnd(Handle);
            using var pen = new Pen(AppTheme.Border, 1);

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            g.DrawRectangle(pen, rect);
        }
    }
}
