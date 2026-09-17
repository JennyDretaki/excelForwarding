using System.ComponentModel;
using ExcelImporter.Models;

namespace ExcelImporter.Services;

public enum PreviewStatus
{
    Ready,
    AlreadyExists,
    Error
}

public sealed class ShortcutPreview
{
    [DisplayName("Δικαστήριο")]
    public string Dikastirio { get; set; } = string.Empty;

    [DisplayName("ΑΦΜ")]
    public string Afm { get; set; } = string.Empty;

    [DisplayName("Host Account")]
    public string HostAccount { get; set; } = string.Empty;

    [DisplayName("Σχόλια")]
    public string Scholia { get; set; } = string.Empty;

    [DisplayName("ΓΑΚ")]
    public string Gak { get; set; } = string.Empty;

    [DisplayName("ΕΑΚ")]
    public string Eak { get; set; } = string.Empty;

    [DisplayName("Ημερομηνία")]
    public string Imerominia { get; set; } = string.Empty;

    [DisplayName("Μήνυμα")]
    public string Message { get; set; } = string.Empty;

    [Browsable(false)]
    public string StatusText { get; set; } = string.Empty;

    [Browsable(false)]
    public PreviewStatus Status { get; set; }

    [Browsable(false)]
    public string Onomateponymo { get; set; } = string.Empty;

    [Browsable(false)]
    public string TargetFolder { get; set; } = string.Empty;

    [Browsable(false)]
    public string ShortcutName { get; set; } = string.Empty;

    [Browsable(false)]
    public string ShortcutPath { get; set; } = string.Empty;

    [Browsable(false)]
    public string SourceFilePath { get; set; } = string.Empty;

    [Browsable(false)]
    public bool CanSend => Status == PreviewStatus.Ready;
}

public sealed class ShortcutService
{
    private readonly FolderLookupService _lookup;
    private readonly ImpersonatedFileService _fileService;

    public ShortcutService(FolderLookupService lookup, ImpersonatedFileService fileService)
    {
        _lookup = lookup;
        _fileService = fileService;
    }

    public IReadOnlyList<ShortcutPreview> BuildPreview(IEnumerable<ExcelRecord> records)
    {
        var previews = new List<ShortcutPreview>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var record in records)
        {
            var preview = new ShortcutPreview
            {
                Dikastirio = record.Dikastirio,
                Afm = record.Afm,
                HostAccount = record.HostAccount,
                Scholia = record.Scholia,
                Gak = record.Gak,
                Eak = record.Eak,
                Imerominia = record.Imerominia,
                Onomateponymo = record.Onomateponymo,
                SourceFilePath = record.SourceFilePath
            };

            try
            {
                if (string.IsNullOrWhiteSpace(record.SourceFilePath) || !File.Exists(record.SourceFilePath))
                {
                    SetError(preview, "Δεν βρέθηκε το αρχικό Excel αρχείο.");
                    previews.Add(preview);
                    continue;
                }

                var match = _lookup.Resolve(record);
                // Συντόμευση .lnk στον φάκελο (όχι αντιγραφή του Excel)
                var shortcutName = Path.GetFileNameWithoutExtension(record.SourceFilePath) + ".lnk";
                var destPath = Path.Combine(match.TargetDirectory, shortcutName);

                preview.TargetFolder = match.TargetDirectory;
                preview.ShortcutName = shortcutName;
                preview.ShortcutPath = destPath;

                if (!seen.Add(destPath))
                {
                    preview.Status = PreviewStatus.AlreadyExists;
                    preview.StatusText = "Διπλότυπο";
                    preview.Message = "Ίδια συντόμευση εμφανίστηκε ήδη στο preview.";
                    previews.Add(preview);
                    continue;
                }

                if (!_fileService.DirectoryExistsWithImpersonate(match.TargetDirectory, out var dirError))
                {
                    SetError(preview, string.IsNullOrWhiteSpace(dirError)
                        ? $"Ο φάκελος προορισμού δεν υπάρχει: {match.TargetDirectory}"
                        : dirError);
                    previews.Add(preview);
                    continue;
                }

                if (_fileService.FileExistsWithImpersonate(destPath, out var existsError))
                {
                    preview.Status = PreviewStatus.AlreadyExists;
                    preview.StatusText = "Υπάρχει ήδη";
                    preview.Message = "Η συντόμευση υπάρχει ήδη στον φάκελο.";
                }
                else if (!string.IsNullOrWhiteSpace(existsError))
                {
                    SetError(preview, existsError);
                }
                else
                {
                    preview.Status = PreviewStatus.Ready;
                    preview.StatusText = "Έτοιμο";
                    preview.Message = "Θα δημιουργηθεί συντόμευση (.lnk) με Αποστολή.";
                }

                previews.Add(preview);
            }
            catch (Exception ex)
            {
                SetError(preview, ex.Message);
                previews.Add(preview);
            }
        }

        return previews;
    }

    public IReadOnlyList<ShortcutResult> Send(IEnumerable<ShortcutPreview> previews)
    {
        var results = new List<ShortcutResult>();

        foreach (var preview in previews.Where(p => p.CanSend))
        {
            try
            {
                if (_fileService.FileExistsWithImpersonate(preview.ShortcutPath, out _))
                {
                    results.Add(new ShortcutResult
                    {
                        HostAccount = preview.HostAccount,
                        Gak = preview.Gak,
                        Onomateponymo = preview.Onomateponymo,
                        TargetPath = preview.ShortcutPath,
                        Action = ShortcutAction.SkippedExisting,
                        Message = "Η συντόμευση υπάρχει ήδη."
                    });
                    continue;
                }

                // Δημιουργία .lnk τοπικά → αποστολή στον φάκελο με impersonation (όπως CBS)
                var tempLnk = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".lnk");
                try
                {
                    CreateLocalShortcut(tempLnk, preview.SourceFilePath);

                    using var ms = ImpersonatedFileService.FileToMemoryStream(tempLnk);
                    if (!_fileService.CopyFileFromStreamWithImpersonate(
                            preview.TargetFolder,
                            ms,
                            preview.ShortcutPath,
                            out var copyError))
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(copyError) ? "Αποτυχία αποστολής συντόμευσης." : copyError);
                    }
                }
                finally
                {
                    if (File.Exists(tempLnk))
                        File.Delete(tempLnk);
                }

                results.Add(new ShortcutResult
                {
                    HostAccount = preview.HostAccount,
                    Gak = preview.Gak,
                    Onomateponymo = preview.Onomateponymo,
                    TargetPath = preview.ShortcutPath,
                    Action = ShortcutAction.Created,
                    Message = "Δημιουργήθηκε συντόμευση."
                });
            }
            catch (Exception ex)
            {
                results.Add(new ShortcutResult
                {
                    HostAccount = preview.HostAccount,
                    Gak = preview.Gak,
                    Onomateponymo = preview.Onomateponymo,
                    TargetPath = preview.ShortcutPath,
                    Action = ShortcutAction.Failed,
                    Message = ex.Message
                });
            }
        }

        return results;
    }

    private static void SetError(ShortcutPreview preview, string message)
    {
        preview.Status = PreviewStatus.Error;
        preview.StatusText = "Σφάλμα";
        preview.Message = message;
    }

    private static void CreateLocalShortcut(string shortcutPath, string targetFile)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("Δεν είναι διαθέσιμο το WScript.Shell για δημιουργία συντόμευσης.");

        dynamic shell = Activator.CreateInstance(shellType)
            ?? throw new InvalidOperationException("Αποτυχία δημιουργίας WScript.Shell.");

        var shortcut = shell.CreateShortcut(shortcutPath);
        shortcut.TargetPath = targetFile;
        shortcut.WorkingDirectory = Path.GetDirectoryName(targetFile) ?? string.Empty;
        shortcut.Description = "Συντόμευση Excel από ExcelImporter";
        shortcut.Save();
    }
}

public enum ShortcutAction
{
    Created,
    SkippedExisting,
    Failed
}

public sealed class ShortcutResult
{
    public string HostAccount { get; init; } = string.Empty;
    public string Gak { get; init; } = string.Empty;
    public string Onomateponymo { get; init; } = string.Empty;
    public string TargetPath { get; init; } = string.Empty;
    public ShortcutAction Action { get; init; }
    public string Message { get; init; } = string.Empty;
}
