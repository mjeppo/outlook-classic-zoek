namespace OutlookClassicSearch;

internal static class AppSettingsParser
{
    public static HashSet<string> ParseExtensions(string raw)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return set;
        }

        var parts = raw.Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            string ext = part.Trim();
            if (string.IsNullOrWhiteSpace(ext))
            {
                continue;
            }

            if (!ext.StartsWith('.'))
            {
                ext = "." + ext;
            }

            set.Add(ext.ToLowerInvariant());
        }

        return set;
    }
}
