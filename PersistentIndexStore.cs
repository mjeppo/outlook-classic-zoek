using System.Text.Json;

namespace OutlookClassicSearch;

internal sealed class PersistentSearchIndex
{
    public DateTime BuiltAtUtc { get; set; }
    public List<string> IncludedFolderEntryIds { get; set; } = new();
    public List<string> IncludedStoreIds { get; set; } = new();
    public List<IndexedMailItem> Items { get; set; } = new();
}

internal sealed class PerStoreSearchIndex
{
    public string StoreId { get; set; } = string.Empty;
    public DateTime BuiltAtUtc { get; set; }
    public List<IndexedMailItem> Items { get; set; } = new();
}

internal static class PersistentIndexStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
    };

    /// <summary>Load the combined index from all per-store files, falling back to the legacy combined file.</summary>
    public static PersistentSearchIndex? Load() => LoadAll();

    public static PersistentSearchIndex? LoadAll()
    {
        try
        {
            var storeFiles = Directory.GetFiles(AppPaths.BaseDirectory, "search-index-store-*.json");
            if (storeFiles.Length > 0)
            {
                var allItems = new List<IndexedMailItem>();
                var allStoreIds = new List<string>();
                var oldestBuilt = DateTime.UtcNow;

                foreach (var file in storeFiles)
                {
                    try
                    {
                        var storeIndex = JsonSerializer.Deserialize<PerStoreSearchIndex>(
                            File.ReadAllText(file), JsonOptions);
                        if (storeIndex is null) continue;
                        allItems.AddRange(storeIndex.Items);
                        if (!allStoreIds.Contains(storeIndex.StoreId))
                            allStoreIds.Add(storeIndex.StoreId);
                        if (storeIndex.BuiltAtUtc < oldestBuilt)
                            oldestBuilt = storeIndex.BuiltAtUtc;
                    }
                    catch { }
                }

                return new PersistentSearchIndex
                {
                    BuiltAtUtc = oldestBuilt,
                    IncludedStoreIds = allStoreIds,
                    Items = allItems
                };
            }

            // Fall back to legacy combined file
            if (!File.Exists(AppPaths.IndexFilePath)) return null;
            return JsonSerializer.Deserialize<PersistentSearchIndex>(
                File.ReadAllText(AppPaths.IndexFilePath), JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    public static PerStoreSearchIndex? LoadForStore(string storeId)
    {
        try
        {
            string path = AppPaths.IndexFilePathForStore(storeId);
            if (!File.Exists(path)) return null;
            return JsonSerializer.Deserialize<PerStoreSearchIndex>(
                File.ReadAllText(path), JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    public static void SaveForStore(string storeId, PerStoreSearchIndex index)
    {
        File.WriteAllText(AppPaths.IndexFilePathForStore(storeId),
            JsonSerializer.Serialize(index, JsonOptions));
    }

    // Legacy combined save — kept so existing callers compile
    public static void Save(PersistentSearchIndex index)
    {
        File.WriteAllText(AppPaths.IndexFilePath,
            JsonSerializer.Serialize(index, JsonOptions));
    }

    public static void Clear() => ClearAll();

    public static void ClearAll()
    {
        foreach (var file in Directory.GetFiles(AppPaths.BaseDirectory, "search-index-store-*.json"))
        {
            try { File.Delete(file); } catch { }
        }
        if (File.Exists(AppPaths.IndexFilePath))
        {
            try { File.Delete(AppPaths.IndexFilePath); } catch { }
        }
    }
}
