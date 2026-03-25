namespace OutlookClassicSearch;

internal static class AppVisualAssets
{
    private static readonly Lazy<Icon?> AppIcon = new(LoadAppIcon);

    public static void ApplyWindowIcon(Form form)
    {
        if (AppIcon.Value is Icon icon)
        {
            form.Icon = (Icon)icon.Clone();
        }
    }

    public static Image? TryLoadLogoImage()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "ocs_full.jpg");
        if (!File.Exists(path))
        {
            return null;
        }

        byte[] bytes = File.ReadAllBytes(path);
        using var ms = new MemoryStream(bytes);
        using var image = Image.FromStream(ms);
        return new Bitmap(image);
    }

    private static Icon? LoadAppIcon()
    {
        try
        {
            return Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
        catch
        {
            return null;
        }
    }
}
