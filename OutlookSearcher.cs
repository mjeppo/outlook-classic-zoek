using System.Runtime.InteropServices;
using Microsoft.Win32;
using Outlook = NetOffice.OutlookApi;

namespace OutlookClassicSearch;

internal sealed class SearchCriteria
{
    public string Query { get; init; } = string.Empty;
    public bool SearchBody { get; init; }
    public bool SearchAttachments { get; init; }
    public bool UsePersistentIndex { get; init; }
    public DateTime? From { get; init; }
    public DateTime? To { get; init; }
    public int MaxResults { get; init; } = 5000;
    public IReadOnlyList<string> SelectedStoreIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ExcludedFolderEntryIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> IncludedFolderPaths { get; init; } = Array.Empty<string>();
    public IReadOnlySet<string> ExcludedAttachmentExtensions { get; init; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
}

internal sealed class MailStoreInfo
{
    public required string DisplayName { get; init; }
    public required string StoreId { get; init; }
}

internal sealed class SearchResultSet
{
    public required IReadOnlyList<EmailSearchResult> Results { get; init; }
    public required int TotalMatches { get; init; }
}

internal sealed class EmailSearchResult
{
    public required string Mailbox { get; init; }
    public required string FolderPath { get; init; }
    public DateTime ReceivedTime { get; init; }
    public required string Subject { get; init; }
    public required string Sender { get; init; }
    public required string Recipients { get; init; }
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public bool HasAttachments { get; init; }
}

internal sealed class MailPreview
{
    public required string Subject { get; init; }
    public required string Sender { get; init; }
    public required string Recipients { get; init; }
    public DateTime ReceivedTime { get; init; }
    public required string Body { get; init; }
}

internal sealed class MailboxFolderRoot
{
    public required string StoreId { get; init; }
    public required string DisplayName { get; init; }
    public required string EntryId { get; init; }
    public required List<MailboxFolderNode> Children { get; init; }
}

internal sealed class MailboxFolderNode
{
    public required string EntryId { get; init; }
    public required string DisplayName { get; init; }
    public required string FolderPath { get; init; }
    public required List<MailboxFolderNode> Children { get; init; }
}

internal sealed class IndexBuildRequest
{
    public IReadOnlyList<string> StoreIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> IncludedFolderEntryIds { get; init; } = Array.Empty<string>();
    public bool SearchBody { get; init; }
    public bool IncludeAttachments { get; init; }
    public IReadOnlySet<string> ExcludedAttachmentExtensions { get; init; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    public bool ForceRebuild { get; init; }
}

internal sealed class IndexedMailItem
{
    public required string Mailbox { get; init; }
    public required string StoreId { get; init; }
    public required string FolderPath { get; init; }
    public required string FolderEntryId { get; init; }
    public required string EntryId { get; init; }
    public required string Subject { get; init; }
    public required string Sender { get; init; }
    public required string ToRecipients { get; init; }
    public required string Body { get; init; }
    public required string AttachmentIndexText { get; init; }
    public DateTime ReceivedTime { get; init; }
    public DateTime LastModificationTime { get; init; }
    public bool HasAttachments { get; init; }
}

internal static class OutlookSearcher
{
    public static Task<IReadOnlyList<MailStoreInfo>> GetAvailableStoresAsync(CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            var stores = new List<MailStoreInfo>();
            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            Outlook.Stores? outlookStores = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                outlookStores = mapi.Stores;

                if (outlookStores is null)
                {
                    throw new InvalidOperationException("Outlook stores konden niet worden opgehaald.");
                }

                for (int i = 1; i <= outlookStores.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    Outlook.Store? store = null;
                    try
                    {
                        store = outlookStores[i] as Outlook.Store;
                        if (store is null)
                        {
                            continue;
                        }

                        stores.Add(new MailStoreInfo
                        {
                            DisplayName = SafeString(store.DisplayName),
                            StoreId = SafeString(store.StoreID)
                        });
                    }
                    finally
                    {
                        ReleaseComObject(store);
                    }
                }

                return (IReadOnlyList<MailStoreInfo>)stores;
            }
            finally
            {
                ReleaseComObject(outlookStores);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, cancellationToken);
    }

    public static Task<IReadOnlyList<MailboxFolderRoot>> GetFolderTreeAsync(
        IReadOnlyList<string> selectedStoreIds,
        CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            var selected = new HashSet<string>(selectedStoreIds, StringComparer.OrdinalIgnoreCase);
            var result = new List<MailboxFolderRoot>();

            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            Outlook.Stores? stores = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                stores = mapi.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    Outlook.Store? store = null;
                    Outlook.MAPIFolder? root = null;
                    try
                    {
                        store = stores[i] as Outlook.Store;
                        if (store is null)
                        {
                            continue;
                        }

                        string storeId = SafeString(store.StoreID);
                        if (selected.Count > 0 && !selected.Contains(storeId))
                        {
                            continue;
                        }

                        root = store.GetRootFolder();
                        if (root is null)
                        {
                            continue;
                        }

                        var rootNode = new MailboxFolderRoot
                        {
                            StoreId = storeId,
                            DisplayName = SafeString(store.DisplayName),
                            EntryId = SafeString(root.EntryID),
                            Children = new List<MailboxFolderNode>()
                        };

                        LoadFolderChildren(root, rootNode.Children, cancellationToken);
                        result.Add(rootNode);
                    }
                    finally
                    {
                        ReleaseComObject(root);
                        ReleaseComObject(store);
                    }
                }

                return (IReadOnlyList<MailboxFolderRoot>)result;
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, cancellationToken);
    }

    public static Task<PersistentSearchIndex> BuildPersistentIndexAsync(
        IndexBuildRequest request,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            var selectedStores = new HashSet<string>(request.StoreIds, StringComparer.OrdinalIgnoreCase);
            var includeSet = new HashSet<string>(request.IncludedFolderEntryIds, StringComparer.OrdinalIgnoreCase);
            var allItems = new List<IndexedMailItem>();

            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            Outlook.Stores? stores = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                stores = mapi.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    Outlook.Store? store = null;
                    Outlook.MAPIFolder? root = null;
                    try
                    {
                        store = stores[i] as Outlook.Store;
                        if (store is null) continue;

                        string storeId = SafeString(store.StoreID);
                        if (selectedStores.Count > 0 && !selectedStores.Contains(storeId)) continue;

                        progress?.Report($"Indexeren mailbox: {store.DisplayName}");
                        root = store.GetRootFolder();
                        if (root is null) continue;

                        // Load existing per-store index for delta comparison (skip if ForceRebuild)
                        var existingStore = request.ForceRebuild ? null : PersistentIndexStore.LoadForStore(storeId);
                        var existingByEntryId = existingStore?.Items
                            .ToDictionary(x => x.EntryId, StringComparer.OrdinalIgnoreCase)
                            ?? new Dictionary<string, IndexedMailItem>(StringComparer.OrdinalIgnoreCase);

                        var storeItems = new List<IndexedMailItem>();
                        int reuseCount = 0;

                        ScanFolders(
                            root,
                            storeId,
                            SafeString(store.DisplayName),
                            includeSet,
                            insideIncluded: false,
                            request.SearchBody,
                            request.IncludeAttachments,
                            request.ExcludedAttachmentExtensions,
                            progress,
                            cancellationToken,
                            (mail, folder) =>
                            {
                                string entryId = SafeString(mail.EntryID);

                                // Try delta: reuse cached entry if item hasn't changed
                                DateTime lastMod = DateTime.MinValue;
                                try { lastMod = mail.LastModificationTime; } catch { }

                                if (lastMod != DateTime.MinValue &&
                                    existingByEntryId.TryGetValue(entryId, out var cached) &&
                                    cached.LastModificationTime == lastMod)
                                {
                                    storeItems.Add(cached);
                                    reuseCount++;
                                    if (storeItems.Count % 50 == 0)
                                        progress?.Report($"{storeItems.Count} items ({reuseCount} ongewijzigd)...");
                                    return false;
                                }

                                // Full index for new or changed items
                                string attachmentIndexText = request.IncludeAttachments
                                    ? BuildAttachmentIndex(mail, request.ExcludedAttachmentExtensions, includeContent: true)
                                    : string.Empty;

                                storeItems.Add(new IndexedMailItem
                                {
                                    Mailbox = SafeString(store.DisplayName),
                                    StoreId = storeId,
                                    FolderPath = SafeString(folder.FolderPath),
                                    FolderEntryId = SafeString(folder.EntryID),
                                    EntryId = entryId,
                                    Subject = SafeString(mail.Subject),
                                    Sender = SafeString(mail.SenderName),
                                    ToRecipients = SafeString(mail.To),
                                    Body = request.SearchBody ? SafeString(mail.Body) : string.Empty,
                                    AttachmentIndexText = attachmentIndexText,
                                    ReceivedTime = mail.ReceivedTime,
                                    LastModificationTime = lastMod,
                                    HasAttachments = GetHasAttachments(mail)
                                });

                                if (storeItems.Count % 50 == 0)
                                    progress?.Report($"{storeItems.Count} items ({reuseCount} ongewijzigd)...");

                                return false;
                            });

                        // Save per-store index immediately after scanning this mailbox
                        PersistentIndexStore.SaveForStore(storeId, new PerStoreSearchIndex
                        {
                            StoreId = storeId,
                            BuiltAtUtc = DateTime.UtcNow,
                            Items = storeItems
                        });

                        allItems.AddRange(storeItems);
                    }
                    finally
                    {
                        ReleaseComObject(root);
                        ReleaseComObject(store);
                    }
                }

                var index = new PersistentSearchIndex
                {
                    BuiltAtUtc = DateTime.UtcNow,
                    IncludedFolderEntryIds = includeSet.ToList(),
                    IncludedStoreIds = selectedStores.ToList(),
                    Items = allItems
                };

                progress?.Report($"Klaar. {allItems.Count} items geindexeerd.");
                return index;
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, cancellationToken);
    }

    public static Task<SearchResultSet> SearchAsync(
        SearchCriteria criteria,
        IProgress<string>? progress,
        IProgress<EmailSearchResult>? resultProgress,
        CancellationToken cancellationToken)
    {
        if (criteria.UsePersistentIndex)
        {
            var persistent = PersistentIndexStore.Load();
            if (persistent is not null)
            {
                var indexedStores = new HashSet<string>(persistent.IncludedStoreIds, StringComparer.OrdinalIgnoreCase);
                bool selectedStoresCovered = criteria.SelectedStoreIds.Count == 0 ||
                    criteria.SelectedStoreIds.All(indexedStores.Contains);

                if (selectedStoresCovered)
                {
                    return Task.FromResult(FilterIndex(persistent, criteria, progress, resultProgress, cancellationToken));
                }

                progress?.Report("Index dekt niet alle geselecteerde mailboxen. Direct zoeken in Outlook...");
            }
        }

        return SearchDirectAsync(criteria, progress, resultProgress, cancellationToken);
    }

    private static SearchResultSet FilterIndex(
        PersistentSearchIndex index,
        SearchCriteria criteria,
        IProgress<string>? progress,
        IProgress<EmailSearchResult>? resultProgress,
        CancellationToken cancellationToken)
    {
        var results = new List<EmailSearchResult>(Math.Min(criteria.MaxResults, 2000));
        var storeSet = new HashSet<string>(criteria.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
        var excludedFolders = new HashSet<string>(criteria.ExcludedFolderEntryIds, StringComparer.OrdinalIgnoreCase);
        int totalMatches = 0;

        var indexAge = DateTime.UtcNow - index.BuiltAtUtc;
        var indexAgeHours = (int)indexAge.TotalHours;

        string ageWarning = indexAgeHours > 24 
            ? $" (⚠ {indexAgeHours / 24} dagen oud - mogelijk verouderd)"
            : indexAgeHours > 6
            ? $" (⚠ {indexAgeHours} uur oud)"
            : "";

        progress?.Report($"Zoeken in index van {index.BuiltAtUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}{ageWarning}");

        // DEBUG: Log aantal items in index en eerste paar subjects
        System.Diagnostics.Debug.WriteLine($"[INDEX] Total items in index: {index.Items.Count}");
        System.Diagnostics.Debug.WriteLine($"[INDEX] Index age: {indexAgeHours} hours");
        System.Diagnostics.Debug.WriteLine($"[INDEX] Query: '{criteria.Query}'");
        int debugCount = 0;
        foreach (var debugItem in index.Items.Take(5))
        {
            System.Diagnostics.Debug.WriteLine($"[INDEX] Sample {++debugCount}: '{debugItem.Subject}'");
        }

        foreach (var item in index.Items)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (storeSet.Count > 0 && !storeSet.Contains(item.StoreId))
            {
                continue;
            }

            if (excludedFolders.Contains(item.FolderEntryId))
            {
                continue;
            }

            if (criteria.IncludedFolderPaths.Count > 0 && !criteria.IncludedFolderPaths.Contains(item.FolderPath))
            {
                continue;
            }

            if (!MatchesDateRange(item.ReceivedTime, criteria.From, criteria.To))
            {
                continue;
            }

            if (!MatchesText(
                    item.Subject,
                    item.Sender,
                    item.ToRecipients,
                    item.Body,
                    criteria.SearchAttachments ? item.AttachmentIndexText : string.Empty,
                    criteria.Query,
                    criteria.SearchBody,
                    criteria.SearchAttachments))
            {
                continue;
            }

            // Tel alle matches
            totalMatches++;

            // Voeg alleen eerste MaxResults toe aan resultatenlijst
            if (results.Count < criteria.MaxResults)
            {
                var result = new EmailSearchResult
                {
                    Mailbox = item.Mailbox,
                    FolderPath = item.FolderPath,
                    ReceivedTime = item.ReceivedTime,
                    Subject = item.Subject,
                    Sender = item.Sender,
                    Recipients = item.ToRecipients,
                    EntryId = item.EntryId,
                    StoreId = item.StoreId,
                    HasAttachments = item.HasAttachments
                };

                results.Add(result);
                resultProgress?.Report(result);
            }
        }

        // Waarschuw als de index mogelijk verouderd is
        if (indexAge.TotalHours > 6)
        {
            var ageText = indexAge.TotalHours > 24 
                ? $"{(int)(indexAge.TotalHours / 24)} dagen"
                : $"{(int)indexAge.TotalHours} uur";

            progress?.Report($"⚠ Index is {ageText} oud. Nieuwe berichten ontbreken mogelijk. Ververs via Instellingen.");
        }

        return new SearchResultSet 
        { 
            Results = results,
            TotalMatches = totalMatches
        };
    }

    private static Task<SearchResultSet> SearchDirectAsync(
        SearchCriteria criteria,
        IProgress<string>? progress,
        IProgress<EmailSearchResult>? resultProgress,
        CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            var results = new List<EmailSearchResult>(Math.Min(criteria.MaxResults, 1000));
            var storeSet = new HashSet<string>(criteria.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
            var excludedFolders = new HashSet<string>(criteria.ExcludedFolderEntryIds, StringComparer.OrdinalIgnoreCase);

            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            Outlook.Stores? stores = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                stores = mapi.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    Outlook.Store? store = null;
                    Outlook.MAPIFolder? root = null;
                    try
                    {
                        store = stores[i] as Outlook.Store;
                        if (store is null)
                        {
                            continue;
                        }

                        string storeId = SafeString(store.StoreID);
                        if (storeSet.Count > 0 && !storeSet.Contains(storeId))
                        {
                            continue;
                        }

                        root = store.GetRootFolder();
                        if (root is null)
                        {
                            continue;
                        }

                        progress?.Report($"Doorzoek mailbox: {store.DisplayName}");

                        bool stop = ScanFolders(
                            root,
                            storeId,
                            SafeString(store.DisplayName),
                            includeSet: null,
                            insideIncluded: false,
                            criteria.SearchBody,
                            criteria.SearchAttachments,
                            criteria.ExcludedAttachmentExtensions,
                            progress,
                            cancellationToken,
                            (mail, folder) =>
                            {
                                if (excludedFolders.Contains(SafeString(folder.EntryID)))
                                {
                                    return false;
                                }

                                string folderPath = SafeString(folder.FolderPath);
                                if (criteria.IncludedFolderPaths.Count > 0 && !criteria.IncludedFolderPaths.Contains(folderPath))
                                {
                                    return false;
                                }

                                if (!MatchesDateRange(mail.ReceivedTime, criteria.From, criteria.To))
                                {
                                    return false;
                                }

                                string attachmentText = criteria.SearchAttachments
                                    ? BuildAttachmentIndex(mail, criteria.ExcludedAttachmentExtensions, includeContent: false)
                                    : string.Empty;

                                if (!MatchesText(
                                        SafeString(mail.Subject),
                                        SafeString(mail.SenderName),
                                        SafeString(mail.To),
                                        criteria.SearchBody ? SafeString(mail.Body) : string.Empty,
                                        attachmentText,
                                        criteria.Query,
                                        criteria.SearchBody,
                                        criteria.SearchAttachments))
                                {
                                    return false;
                                }

                                var result = new EmailSearchResult
                                {
                                    Mailbox = SafeString(store.DisplayName),
                                    FolderPath = SafeString(folder.FolderPath),
                                    ReceivedTime = mail.ReceivedTime,
                                    Subject = SafeString(mail.Subject),
                                    Sender = SafeString(mail.SenderName),
                                    Recipients = SafeString(mail.To),
                                    EntryId = SafeString(mail.EntryID),
                                    StoreId = storeId,
                                    HasAttachments = GetHasAttachments(mail)
                                };

                                results.Add(result);
                                resultProgress?.Report(result);
                                return results.Count >= criteria.MaxResults;
                            });

                        if (stop || results.Count >= criteria.MaxResults)
                        {
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(root);
                        ReleaseComObject(store);
                    }
                }

                return new SearchResultSet
                {
                    Results = results,
                    TotalMatches = results.Count // Voor DirectSearch kunnen we niet verder tellen zonder alles te scannen
                };
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, cancellationToken);
    }

    private static bool ScanFolders(
        Outlook.MAPIFolder folder,
        string storeId,
        string mailbox,
        HashSet<string>? includeSet,
        bool insideIncluded,
        bool searchBody,
        bool searchAttachments,
        IReadOnlySet<string> excludedAttachmentExtensions,
        IProgress<string>? progress,
        CancellationToken cancellationToken,
        Func<Outlook.MailItem, Outlook.MAPIFolder, bool> onMailFound)
    {
        cancellationToken.ThrowIfCancellationRequested();

        string currentEntryId = SafeString(folder.EntryID);
        // A folder is included if: no includeSet, OR we're inside an already-included subtree,
        // OR this folder is explicitly in the includeSet.
        bool includeFolder = insideIncluded || includeSet is null || includeSet.Count == 0 || includeSet.Contains(currentEntryId);

        Outlook.Items? items = null;
        Outlook.Folders? children = null;

        try
        {
            progress?.Report($"Map: {SafeString(folder.FolderPath)}");

            if (includeFolder)
            {
                items = folder.Items as Outlook.Items;
                if (items is not null)
                {
                    int count = items.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        object? raw = null;
                        try
                        {
                            raw = items[i];
                            if (raw is not Outlook.MailItem mail)
                            {
                                continue;
                            }

                            if (onMailFound(mail, folder))
                            {
                                return true;
                            }
                        }
                        finally
                        {
                            ReleaseComObject(raw);
                        }
                    }
                }
            }

            children = folder.Folders as Outlook.Folders;
            if (children is null)
            {
                return false;
            }

            int childCount = children.Count;
            for (int i = 1; i <= childCount; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                Outlook.MAPIFolder? child = null;
                try
                {
                    child = children[i] as Outlook.MAPIFolder;
                    if (child is null)
                    {
                        continue;
                    }

                    if (ScanFolders(
                        child,
                        storeId,
                        mailbox,
                        includeSet,
                        insideIncluded || includeFolder,
                        searchBody,
                        searchAttachments,
                        excludedAttachmentExtensions,
                        progress,
                        cancellationToken,
                        onMailFound))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Skip inaccessible folders.
                }
                finally
                {
                    ReleaseComObject(child);
                }
            }
        }
        finally
        {
            ReleaseComObject(children);
            ReleaseComObject(items);
        }

        return false;
    }

    private static string BuildAttachmentIndex(
        Outlook.MailItem mail,
        IReadOnlySet<string> excludedExtensions,
        bool includeContent)
    {
        var tokens = new List<string>();
        Outlook.Attachments? attachments = null;

        try
        {
            attachments = mail.Attachments as Outlook.Attachments;
            if (attachments is null)
            {
                return string.Empty;
            }

            int count = attachments.Count;
            for (int i = 1; i <= count; i++)
            {
                Outlook.Attachment? attachment = null;
                try
                {
                    attachment = attachments[i] as Outlook.Attachment;
                    if (attachment is null)
                    {
                        continue;
                    }

                    string name = SafeString(attachment.FileName);
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    string ext = Path.GetExtension(name).ToLowerInvariant();
                    if (excludedExtensions.Contains(ext))
                    {
                        continue;
                    }

                    tokens.Add(name);

                    if (includeContent && IsTextLikeExtension(ext))
                    {
                        string temp = Path.Combine(Path.GetTempPath(), $"ocs-{Guid.NewGuid():N}{ext}");
                        try
                        {
                            attachment.SaveAsFile(temp);
                            string text = File.ReadAllText(temp);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                tokens.Add(text);
                            }
                        }
                        catch
                        {
                            // Ignore unreadable attachment content.
                        }
                        finally
                        {
                            try
                            {
                                if (File.Exists(temp))
                                {
                                    File.Delete(temp);
                                }
                            }
                            catch
                            {
                                // Ignore temp cleanup issues.
                            }
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(attachment);
                }
            }
        }
        finally
        {
            ReleaseComObject(attachments);
        }

        return string.Join(" ", tokens);
    }

    private static bool IsTextLikeExtension(string ext)
    {
        return ext is ".txt" or ".csv" or ".log" or ".json" or ".xml";
    }

    private static bool MatchesText(
        string subject,
        string sender,
        string toRecipients,
        string body,
        string attachmentText,
        string query,
        bool searchBody,
        bool searchAttachments)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        subject ??= string.Empty;
        sender ??= string.Empty;
        toRecipients ??= string.Empty;
        body ??= string.Empty;
        attachmentText ??= string.Empty;

        string normalizedQuery = query.Trim();
        bool usesTokenQuery = normalizedQuery.Contains('+') || normalizedQuery.Contains('-');

        if (usesTokenQuery)
        {
            string searchable = string.Join(
                "\n",
                subject,
                sender,
                toRecipients,
                searchBody ? body : string.Empty,
                searchAttachments ? attachmentText : string.Empty);

            return MatchesTokenQuery(searchable, normalizedQuery);
        }

        var comparison = StringComparison.OrdinalIgnoreCase;

        // Exact match first (fastest)
        if (subject.Contains(query, comparison) ||
            sender.Contains(query, comparison) ||
            toRecipients.Contains(query, comparison))
        {
            return true;
        }

        if (searchBody && body.Contains(query, comparison))
        {
            return true;
        }

        if (searchAttachments && attachmentText.Contains(query, comparison))
        {
            return true;
        }

        // Fuzzy match for queries with 5+ chars (tolerates 1-2 typos)
        if (query.Length >= 5)
        {
            if (ContainsFuzzy(subject, query, comparison) ||
                ContainsFuzzy(sender, query, comparison) ||
                ContainsFuzzy(toRecipients, query, comparison))
            {
                return true;
            }

            if (searchBody && ContainsFuzzy(body, query, comparison))
            {
                return true;
            }

            if (searchAttachments && ContainsFuzzy(attachmentText, query, comparison))
            {
                return true;
            }
        }

        return false;
    }

    private static bool MatchesTokenQuery(string searchableText, string query)
    {
        string haystack = searchableText;
        var required = new List<string>();
        var excluded = new List<string>();
        var optional = new List<string>();

        foreach (string raw in query.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (raw.StartsWith('+') && raw.Length > 1)
            {
                required.Add(raw[1..]);
                continue;
            }

            if (raw.StartsWith('-') && raw.Length > 1)
            {
                excluded.Add(raw[1..]);
                continue;
            }

            optional.Add(raw);
        }

        var comparison = StringComparison.OrdinalIgnoreCase;

        foreach (string token in excluded)
        {
            // For excluded tokens, use exact match or fuzzy for longer tokens
            if (haystack.Contains(token, comparison) || 
                (token.Length >= 5 && ContainsFuzzy(haystack, token, comparison)))
            {
                return false;
            }
        }

        foreach (string token in required)
        {
            // For required tokens, use exact match or fuzzy for longer tokens
            if (!haystack.Contains(token, comparison) && 
                !(token.Length >= 5 && ContainsFuzzy(haystack, token, comparison)))
            {
                return false;
            }
        }

        if (optional.Count == 0)
        {
            return required.Count > 0;
        }

        foreach (string token in optional)
        {
            // For optional tokens, use exact match or fuzzy for longer tokens
            if (haystack.Contains(token, comparison) || 
                (token.Length >= 5 && ContainsFuzzy(haystack, token, comparison)))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Check if text contains query with fuzzy matching (allows 1-2 character differences).
    /// Splits text into words and checks if any word is similar to the query.
    /// </summary>
    private static bool ContainsFuzzy(string text, string query, StringComparison comparison)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(query))
        {
            return false;
        }

        // Split text into words and check each word
        var words = text.Split(new[] { ' ', '\t', '\n', '\r', '.', ',', ';', ':', '!', '?' }, 
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        int maxDistance = query.Length <= 7 ? 1 : 2; // Allow 1 typo for short words, 2 for longer

        foreach (var word in words)
        {
            if (word.Length < 3) continue; // Skip very short words

            // Quick length check - if difference is > maxDistance, skip expensive calculation
            int lengthDiff = Math.Abs(word.Length - query.Length);
            if (lengthDiff > maxDistance)
            {
                continue;
            }

            int distance = LevenshteinDistance(word, query, comparison);

            if (distance <= maxDistance)
            {
                System.Diagnostics.Debug.WriteLine($"[FUZZY] ✓ MATCH: '{word}' ≈ '{query}' (distance={distance})");
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Calculate Levenshtein distance between two strings (number of edits needed).
    /// Optimized version that stops early if distance exceeds maxDistance.
    /// </summary>
    private static int LevenshteinDistance(string s1, string s2, StringComparison comparison)
    {
        // Normalize for case-insensitive comparison
        if (comparison == StringComparison.OrdinalIgnoreCase)
        {
            s1 = s1.ToLowerInvariant();
            s2 = s2.ToLowerInvariant();
        }

        int len1 = s1.Length;
        int len2 = s2.Length;

        // Quick check for exact match
        if (s1 == s2) return 0;

        // If one string is empty, distance is length of other
        if (len1 == 0) return len2;
        if (len2 == 0) return len1;

        // Create distance matrix (only need current and previous row)
        int[] previousRow = new int[len2 + 1];
        int[] currentRow = new int[len2 + 1];

        // Initialize first row
        for (int j = 0; j <= len2; j++)
        {
            previousRow[j] = j;
        }

        // Calculate distances
        for (int i = 1; i <= len1; i++)
        {
            currentRow[0] = i;

            for (int j = 1; j <= len2; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;

                currentRow[j] = Math.Min(
                    Math.Min(
                        currentRow[j - 1] + 1,      // insertion
                        previousRow[j] + 1),        // deletion
                    previousRow[j - 1] + cost);     // substitution
            }

            // Swap rows
            var temp = previousRow;
            previousRow = currentRow;
            currentRow = temp;
        }

        return previousRow[len2];
    }

    private static bool MatchesDateRange(DateTime receivedTime, DateTime? from, DateTime? to)
    {
        if (from.HasValue && receivedTime < from.Value)
        {
            return false;
        }

        if (to.HasValue && receivedTime > to.Value)
        {
            return false;
        }

        return true;
    }

    public static Task OpenMailItemAsync(string entryId, string storeId)
    {
        return RunOnStaThreadAsync(() =>
        {
            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            object? item = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                item = mapi.GetItemFromID(entryId, storeId);

                if (item is Outlook.MailItem mailItem)
                {
                    mailItem.Display(false);
                }

                return true;
            }
            finally
            {
                ReleaseComObject(item);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, CancellationToken.None);
    }

    public static Task<MailPreview?> GetMailPreviewAsync(string entryId, string storeId, CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            Outlook.Application? app = null;
            Outlook.NameSpace? mapi = null;
            object? item = null;

            try
            {
                (app, mapi) = CreateOutlookSession();
                item = mapi.GetItemFromID(entryId, storeId);

                if (item is not Outlook.MailItem mail)
                {
                    return null;
                }

                return new MailPreview
                {
                    Subject = SafeString(mail.Subject),
                    Sender = SafeString(mail.SenderName),
                    Recipients = SafeString(mail.To),
                    ReceivedTime = mail.ReceivedTime,
                    Body = SafeString(mail.Body)
                };
            }
            finally
            {
                ReleaseComObject(item);
                ReleaseComObject(mapi);
                ReleaseComObject(app);
            }
        }, cancellationToken);
    }

    private static void LoadFolderChildren(Outlook.MAPIFolder parent, List<MailboxFolderNode> output, CancellationToken cancellationToken)
    {
        Outlook.Folders? children = null;
        try
        {
            children = parent.Folders as Outlook.Folders;
            if (children is null)
            {
                return;
            }

            int count = children.Count;
            for (int i = 1; i <= count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                Outlook.MAPIFolder? child = null;
                try
                {
                    child = children[i] as Outlook.MAPIFolder;
                    if (child is null)
                    {
                        continue;
                    }

                    var node = new MailboxFolderNode
                    {
                        EntryId = SafeString(child.EntryID),
                        DisplayName = SafeString(child.Name),
                        FolderPath = SafeString(child.FolderPath),
                        Children = new List<MailboxFolderNode>()
                    };

                    LoadFolderChildren(child, node.Children, cancellationToken);
                    output.Add(node);
                }
                catch
                {
                    // Skip inaccessible folders.
                }
                finally
                {
                    ReleaseComObject(child);
                }
            }
        }
        finally
        {
            ReleaseComObject(children);
        }
    }

    private static bool GetHasAttachments(Outlook.MailItem mail)
    {
        // Quick check via PR_HASATTACH (0x0E1B000B) — avoids iterating attachments when there are none
        try
        {
            const string prHasAttach = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
            var val = mail.PropertyAccessor.GetProperty(prHasAttach);
            if (val is bool b && !b) return false;
            if (val is int iv && iv == 0) return false;
        }
        catch { /* Property not available, proceed with detailed check */ }

        Outlook.Attachments? attachments = null;
        try
        {
            attachments = mail.Attachments as Outlook.Attachments;
            if (attachments is null || attachments.Count == 0) return false;

            int count = attachments.Count;
            for (int i = 1; i <= count; i++)
            {
                Outlook.Attachment? att = null;
                try
                {
                    att = attachments[i] as Outlook.Attachment;
                    if (att is null) continue;

                    // Stap 1: type-filter
                    // olByValue (1) = echte bijlage, olByReference (4) = gekoppeld bestand
                    // olEmbeddeditem (5) = ingesloten origineel bericht → overslaan
                    // olOLE (6) = ingesloten OLE-object → overslaan
                    var attType = att.Type;
                    if (attType != NetOffice.OutlookApi.Enums.OlAttachmentType.olByValue &&
                        attType != NetOffice.OutlookApi.Enums.OlAttachmentType.olByReference)
                        continue;

                    // Stap 2: bestandsnaam-check — lege naam = geen echte bijlage
                    string name = string.Empty;
                    try { name = att.FileName ?? string.Empty; } catch { }
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    // Stap 3: PR_ATTACH_FLAGS check
                    // Bit 4 (ATT_MHTML_REF = 4) = inline afbeelding in HTML-body
                    // (bijv. logo in e-mailhandtekening) → overslaan
                    try
                    {
                        const string prAttachFlags = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
                        var raw = att.PropertyAccessor.GetProperty(prAttachFlags);
                        int flags = raw is int i2 ? i2 : Convert.ToInt32(raw);
                        if ((flags & 4) != 0) continue;
                    }
                    catch { /* property niet aanwezig = echte bijlage, doorgaan */ }

                    return true;
                }
                catch { /* sla bijlage over bij COM-fout */ }
                finally
                {
                    ReleaseComObject(att);
                }
            }

            return false;
        }
        catch { return false; }
        finally
        {
            ReleaseComObject(attachments);
        }
    }

    private static T WithOutlookSession<T>(Func<Outlook.Application, Outlook.NameSpace, T> action)    {
        Outlook.Application? app = null;
        Outlook.NameSpace? mapi = null;
        try
        {
            (app, mapi) = CreateOutlookSession();
            return action(app, mapi);
        }
        finally
        {
            ReleaseComObject(mapi);
            ReleaseComObject(app);
        }
    }

    private static (Outlook.Application app, Outlook.NameSpace mapi) CreateOutlookSession()
    {
        Outlook.Application? app = null;
        Outlook.NameSpace? mapi = null;

        try
        {
            app = new Outlook.Application();
            mapi = (app.Session ?? app.GetNamespace("MAPI")) as Outlook.NameSpace;

            if (mapi is null)
            {
                throw new InvalidOperationException("MAPI sessie kon niet worden gestart.");
            }

            _ = mapi.CurrentUser?.Name;
            return (app, mapi);
        }
        catch (FileNotFoundException ex)
        {
            ReleaseComObject(mapi);
            ReleaseComObject(app);

            string bitnessHint = BuildOutlookBitnessHint();
            throw new InvalidOperationException(
                "Outlook COM kan niet worden geladen (FileNotFoundException). " +
                "Dit komt meestal door een onjuiste architectuur-build of ontbrekende Office-registratie. " +
                bitnessHint,
                ex);
        }
        catch
        {
            ReleaseComObject(mapi);
            ReleaseComObject(app);
            throw;
        }
    }

    private static string BuildOutlookBitnessHint()
    {
        string appBitness = Environment.Is64BitProcess ? "x64" : "x86";
        string outlookBitness = DetectOutlookBitness();

        if (!string.Equals(outlookBitness, "Onbekend", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(outlookBitness, appBitness, StringComparison.OrdinalIgnoreCase))
        {
            return $"Gedetecteerd: app={appBitness}, Outlook={outlookBitness}. Gebruik de {outlookBitness}-publish map.";
        }

        return $"Gedetecteerd: app={appBitness}, Outlook={outlookBitness}. Probeer de win-x86 build als Outlook 32-bit is, anders win-x64.";
    }

    private static string DetectOutlookBitness()
    {
        string? c2r = ReadRegistryValue(
            RegistryHive.LocalMachine,
            @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            "Platform");

        if (!string.IsNullOrWhiteSpace(c2r))
        {
            return NormalizeBitness(c2r);
        }

        foreach (string version in new[] { "16.0", "15.0" })
        {
            string? value = ReadRegistryValue(
                RegistryHive.LocalMachine,
                $@"SOFTWARE\Microsoft\Office\{version}\Outlook",
                "Bitness");

            if (!string.IsNullOrWhiteSpace(value))
            {
                return NormalizeBitness(value);
            }
        }

        return "Onbekend";
    }

    private static string? ReadRegistryValue(RegistryHive hive, string subKey, string valueName)
    {
        foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            try
            {
                using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view);
                using RegistryKey? key = baseKey.OpenSubKey(subKey, false);
                object? value = key?.GetValue(valueName);
                if (value is string text && !string.IsNullOrWhiteSpace(text))
                {
                    return text;
                }
            }
            catch
            {
                // Ignore registry read issues and continue.
            }
        }

        return null;
    }

    private static string NormalizeBitness(string bitness)
    {
        string normalized = bitness.Trim().ToLowerInvariant();
        if (normalized.Contains("64"))
        {
            return "x64";
        }

        if (normalized.Contains("86") || normalized == "x32" || normalized == "32")
        {
            return "x86";
        }

        return "Onbekend";
    }

    private static string SafeString(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
    }

    private static async Task<T> RunOnStaThreadAsync<T>(Func<T> function, CancellationToken cancellationToken)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }

                T result = function();
                tcs.TrySetResult(result);
            }
            catch (OperationCanceledException oce)
            {
                tcs.TrySetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });

        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        using var registration = cancellationToken.Register(() => tcs.TrySetCanceled(cancellationToken));
        return await tcs.Task.ConfigureAwait(false);
    }

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        if (comObject is IDisposable disposable)
        {
            disposable.Dispose();
            return;
        }

        if (Marshal.IsComObject(comObject))
        {
            Marshal.FinalReleaseComObject(comObject);
        }
    }
}
