using System.Text.Json;
using System.Windows.Forms;

namespace OutlookClassicSearch;

internal sealed class AppSettings
{
    public List<string> SelectedStoreIds { get; set; } = new();
    public List<string> SelectedStoreDisplayNames { get; set; } = new();
    public List<string> ExcludedFolderEntryIds { get; set; } = new();
    public bool SearchBody { get; set; } = true;
    public bool SearchAttachments { get; set; } = false;
    public string ExcludedAttachmentExtensionsRaw { get; set; } = ".zip;.png;.jpg;.jpeg;.gif;.mp4;.mov";
    public bool UseDateRange { get; set; } = true;
    public DateTime DateFrom { get; set; } = DateTime.Today.AddYears(-1);
    public DateTime DateTo { get; set; } = DateTime.Today;
    public int MaxResults { get; set; } = 500;
    public bool UsePersistentIndexForSearch { get; set; } = true;
    public bool ExcludeAttachmentExtensions { get; set; } = true;

    public List<string> IndexIncludedFolderEntryIds { get; set; } = new();
    public bool IndexAutoRefreshEnabled { get; set; } = false;
    public int IndexRefreshIntervalMinutes { get; set; } = 60;
    public string Language { get; set; } = "nl";
    public int SearchHistoryMaxCount { get; set; } = 10;
    public List<string> SearchHistory { get; set; } = new();

    // Venstergrootte en -positie
    public FormWindowState WindowState { get; set; } = FormWindowState.Normal;
    public int WindowLeft { get; set; } = -1;
    public int WindowTop { get; set; } = -1;
    public int WindowWidth { get; set; } = -1;
    public int WindowHeight { get; set; } = -1;
}

internal static class AppSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static AppSettings Load()
    {
        try
        {
            if (!File.Exists(AppPaths.SettingsFilePath))
            {
                return new AppSettings();
            }

            string json = File.ReadAllText(AppPaths.SettingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        string json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(AppPaths.SettingsFilePath, json);
    }
}
