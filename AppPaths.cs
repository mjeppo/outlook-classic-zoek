namespace OutlookClassicSearch;

internal static class AppPaths
{
    public static string BaseDirectory
    {
        get
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookClassicSearch");
            Directory.CreateDirectory(path);
            return path;
        }
    }

    public static string SettingsFilePath => Path.Combine(BaseDirectory, "settings.json");

    public static string IndexFilePath => Path.Combine(BaseDirectory, "search-index.json");

    public static string IndexFilePathForStore(string storeId)
    {
        var hash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(
                System.Text.Encoding.UTF8.GetBytes(storeId)))[..16];
        return Path.Combine(BaseDirectory, $"search-index-store-{hash}.json");
    }
}
