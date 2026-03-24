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
    public int MaxResults { get; init; } = 500;
    public IReadOnlyList<string> SelectedStoreIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ExcludedFolderEntryIds { get; init; } = Array.Empty<string>();
    public IReadOnlySet<string> ExcludedAttachmentExtensions { get; init; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
}

internal sealed class MailStoreInfo
{
    public required string DisplayName { get; init; }
    public required string StoreId { get; init; }
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
            var items = new List<IndexedMailItem>();

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
                        if (selectedStores.Count > 0 && !selectedStores.Contains(storeId))
                        {
                            continue;
                        }

                        progress?.Report($"Indexeren mailbox: {store.DisplayName}");
                        root = store.GetRootFolder();
                        if (root is null)
                        {
                            continue;
                        }

                        ScanFolders(
                            root,
                            storeId,
                            SafeString(store.DisplayName),
                            includeSet,
                            request.SearchBody,
                            request.IncludeAttachments,
                            request.ExcludedAttachmentExtensions,
                            progress,
                            cancellationToken,
                            (mail, folder) =>
                            {
                                string attachmentIndexText = request.IncludeAttachments
                                    ? BuildAttachmentIndex(mail, request.ExcludedAttachmentExtensions, includeContent: true)
                                    : string.Empty;

                                items.Add(new IndexedMailItem
                                {
                                    Mailbox = SafeString(store.DisplayName),
                                    StoreId = storeId,
                                    FolderPath = SafeString(folder.FolderPath),
                                    FolderEntryId = SafeString(folder.EntryID),
                                    EntryId = SafeString(mail.EntryID),
                                    Subject = SafeString(mail.Subject),
                                    Sender = SafeString(mail.SenderName),
                                    ToRecipients = SafeString(mail.To),
                                    Body = request.SearchBody ? SafeString(mail.Body) : string.Empty,
                                    AttachmentIndexText = attachmentIndexText,
                                    ReceivedTime = mail.ReceivedTime
                                });

                                return false;
                            });
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
                    Items = items
                };

                PersistentIndexStore.Save(index);
                progress?.Report($"Klaar. {items.Count} items geindexeerd.");
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

    public static Task<IReadOnlyList<EmailSearchResult>> SearchAsync(
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

    private static IReadOnlyList<EmailSearchResult> FilterIndex(
        PersistentSearchIndex index,
        SearchCriteria criteria,
        IProgress<string>? progress,
        IProgress<EmailSearchResult>? resultProgress,
        CancellationToken cancellationToken)
    {
        var results = new List<EmailSearchResult>(Math.Min(criteria.MaxResults, 2000));
        var storeSet = new HashSet<string>(criteria.SelectedStoreIds, StringComparer.OrdinalIgnoreCase);
        var excludedFolders = new HashSet<string>(criteria.ExcludedFolderEntryIds, StringComparer.OrdinalIgnoreCase);

        progress?.Report($"Zoeken in index van {index.BuiltAtUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}");

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

            var result = new EmailSearchResult
            {
                Mailbox = item.Mailbox,
                FolderPath = item.FolderPath,
                ReceivedTime = item.ReceivedTime,
                Subject = item.Subject,
                Sender = item.Sender,
                Recipients = item.ToRecipients,
                EntryId = item.EntryId,
                StoreId = item.StoreId
            };

            results.Add(result);
            resultProgress?.Report(result);

            if (results.Count >= criteria.MaxResults)
            {
                break;
            }
        }

        return results;
    }

    private static Task<IReadOnlyList<EmailSearchResult>> SearchDirectAsync(
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
                                    StoreId = storeId
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

                return (IReadOnlyList<EmailSearchResult>)results;
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
        bool searchBody,
        bool searchAttachments,
        IReadOnlySet<string> excludedAttachmentExtensions,
        IProgress<string>? progress,
        CancellationToken cancellationToken,
        Func<Outlook.MailItem, Outlook.MAPIFolder, bool> onMailFound)
    {
        cancellationToken.ThrowIfCancellationRequested();

        string currentEntryId = SafeString(folder.EntryID);
        bool includeFolder = includeSet is null || includeSet.Count == 0 || includeSet.Contains(currentEntryId);

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
        if (normalizedQuery.Contains('+') || normalizedQuery.Contains('-'))
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
            if (haystack.Contains(token, comparison))
            {
                return false;
            }
        }

        foreach (string token in required)
        {
            if (!haystack.Contains(token, comparison))
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
            if (haystack.Contains(token, comparison))
            {
                return true;
            }
        }

        return false;
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
