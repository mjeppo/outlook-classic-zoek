using System.Runtime.InteropServices;
using System.Text;
using MaterialSkin;
using MaterialSkin.Controls;

namespace OutlookClassicSearch;

internal static class ErrorDetailsDialog
{
    public static void Show(IWin32Window owner, string title, string summary, Exception exception, string? hint = null)
    {
        string details = BuildDetails(exception);
        string logPath = WriteLog(title, summary, details);

        string message =
            $"{summary}\n\n" +
            $"{(string.IsNullOrWhiteSpace(hint) ? string.Empty : hint + "\n\n")}" +
            $"Logbestand:\n{logPath}\n\n" +
            "Klik op OK voor technische details.";

        MessageBox.Show(owner, message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        using var form = new MaterialForm();
        var materialSkinManager = MaterialSkinManager.Instance;
        materialSkinManager.AddFormToManage(form);

        using var textBox = new TextBox();
        using var copyButton = new Button();

        form.Text = $"{title} - Technische details";
        form.StartPosition = FormStartPosition.CenterParent;
        form.Size = new Size(980, 620);
        form.MinimumSize = new Size(780, 420);
        AppVisualAssets.ApplyWindowIcon(form);

        textBox.Multiline = true;
        textBox.ReadOnly = true;
        textBox.ScrollBars = ScrollBars.Both;
        textBox.WordWrap = false;
        textBox.Dock = DockStyle.Fill;
        textBox.Font = new Font("Consolas", 10);
        textBox.Text = details;

        copyButton.Text = "Kopieer details";
        copyButton.Height = 34;
        copyButton.Dock = DockStyle.Bottom;
        copyButton.Click += (_, _) =>
        {
            Clipboard.SetText(details);
            MessageBox.Show(form, "Details zijn gekopieerd naar klembord.", "Gekopieerd", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };

        form.Controls.Add(textBox);
        form.Controls.Add(copyButton);
        AppTheme.Apply(form);
        AppTheme.ApplyPrimaryStyle(copyButton);
        form.ShowDialog(owner);
    }

    private static string WriteLog(string title, string summary, string details)
    {
        string baseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookClassicSearch",
            "logs");

        Directory.CreateDirectory(baseDir);

        string logPath = Path.Combine(baseDir, $"error-{DateTime.Now:yyyyMMdd}.log");
        var lines = new[]
        {
            "============================================================",
            $"Timestamp : {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            $"Title     : {title}",
            $"Summary   : {summary}",
            details,
            string.Empty
        };

        File.AppendAllLines(logPath, lines);
        return logPath;
    }

    private static string BuildDetails(Exception ex)
    {
        var sb = new StringBuilder();

        sb.AppendLine("Outlook Classic Search - Error details");
        sb.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"OS: {Environment.OSVersion}");
        sb.AppendLine($"64-bit process: {Environment.Is64BitProcess}");
        sb.AppendLine($"CLR: {Environment.Version}");
        sb.AppendLine();

        int level = 0;
        Exception? current = ex;
        while (current is not null)
        {
            sb.AppendLine($"Exception level {level}: {current.GetType().FullName}");
            sb.AppendLine($"Message: {current.Message}");
            sb.AppendLine($"HResult: 0x{current.HResult:X8}");

            if (current is COMException com)
            {
                sb.AppendLine($"COM ErrorCode: 0x{com.ErrorCode:X8}");
            }

            if (!string.IsNullOrWhiteSpace(current.StackTrace))
            {
                sb.AppendLine("StackTrace:");
                sb.AppendLine(current.StackTrace);
            }

            sb.AppendLine();
            current = current.InnerException;
            level++;
        }

        return sb.ToString();
    }
}
