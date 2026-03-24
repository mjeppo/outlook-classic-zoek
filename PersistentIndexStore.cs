using System.Text.Json;

namespace OutlookClassicSearch;

internal sealed class PersistentSearchIndex
{
    public DateTime BuiltAtUtc { get; set; }
    public List<string> IncludedFolderEntryIds { get; set; } = new();
    public List<string> IncludedStoreIds { get; set; } = new();
    public List<IndexedMailItem> Items { get; set; } = new();
}

internal static class PersistentIndexStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    public static PersistentSearchIndex? Load()
    {
        try
        {
            if (!File.Exists(AppPaths.IndexFilePath))
            {
                return null;
            }

            string json = File.ReadAllText(AppPaths.IndexFilePath);
            return JsonSerializer.Deserialize<PersistentSearchIndex>(json, JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    public static void Save(PersistentSearchIndex index)
    {
        string json = JsonSerializer.Serialize(index, JsonOptions);
        File.WriteAllText(AppPaths.IndexFilePath, json);
    }

    public static void Clear()
    {
        if (File.Exists(AppPaths.IndexFilePath))
        {
            File.Delete(AppPaths.IndexFilePath);
        }
    }
}
