using System.Diagnostics;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookClassicSearch;

/// <summary>
/// Controleert op updates via GitHub Releases API
/// </summary>
internal static class UpdateChecker
{
    private const string GitHubApiUrl = "https://api.github.com/repos/mjeppo/outlook-classic-zoek/releases/latest";
    private const string GitHubReleasesUrl = "https://github.com/mjeppo/outlook-classic-zoek/releases";
    private static readonly HttpClient _httpClient = new();

    static UpdateChecker()
    {
        // GitHub API vereist User-Agent header
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "OutlookClassicSearch");
    }

    /// <summary>
    /// Controleert of er een nieuwere versie beschikbaar is
    /// </summary>
    public static async Task<UpdateInfo?> CheckForUpdateAsync()
    {
        try
        {
            var response = await _httpClient.GetStringAsync(GitHubApiUrl);
            var release = JsonSerializer.Deserialize<GitHubRelease>(response);

            if (release == null || string.IsNullOrEmpty(release.TagName))
                return null;

            // Parse versies (verwacht formaat: v1.4.0 of 1.4.0)
            var latestVersion = ParseVersion(release.TagName);
            var currentVersion = ParseVersion(GetCurrentVersion());

            if (latestVersion > currentVersion)
            {
                return new UpdateInfo
                {
                    LatestVersion = release.TagName.TrimStart('v'),
                    CurrentVersion = GetCurrentVersion(),
                    ReleaseUrl = release.HtmlUrl ?? GitHubReleasesUrl,
                    ReleaseNotes = release.Body ?? "",
                    PublishedAt = release.PublishedAt
                };
            }

            return null; // Geen update beschikbaar
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Update check failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Opent de GitHub releases pagina in de browser
    /// </summary>
    public static void OpenReleasesPage()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = GitHubReleasesUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to open releases page: {ex.Message}");
        }
    }

    /// <summary>
    /// Haalt de huidige applicatie versie op
    /// </summary>
    private static string GetCurrentVersion()
    {
        var version = typeof(UpdateChecker).Assembly.GetName().Version;
        return $"{version?.Major}.{version?.Minor}.{version?.Build}";
    }

    /// <summary>
    /// Parse een versie string naar een vergelijkbare Version object
    /// </summary>
    private static Version ParseVersion(string versionString)
    {
        // Verwijder 'v' prefix als aanwezig
        versionString = versionString.TrimStart('v');

        // Probeer te parsen
        if (Version.TryParse(versionString, out var version))
            return version;

        // Fallback naar 0.0.0
        return new Version(0, 0, 0);
    }

    /// <summary>
    /// GitHub Release API response model
    /// </summary>
    private class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string? TagName { get; set; }

        [JsonPropertyName("html_url")]
        public string? HtmlUrl { get; set; }

        [JsonPropertyName("body")]
        public string? Body { get; set; }

        [JsonPropertyName("published_at")]
        public DateTime PublishedAt { get; set; }
    }
}

/// <summary>
/// Informatie over een beschikbare update
/// </summary>
internal class UpdateInfo
{
    public required string LatestVersion { get; init; }
    public required string CurrentVersion { get; init; }
    public required string ReleaseUrl { get; init; }
    public required string ReleaseNotes { get; init; }
    public DateTime PublishedAt { get; init; }
}
