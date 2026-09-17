using System.ComponentModel;
using ExcelImporter.Models;
using ExcelImporter.Services;

namespace ExcelImporter;

public partial class Form1 : Form
{
    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xls",
        ".xlsx"
    };

    private static readonly Color DropIdleColor = Color.FromArgb(245, 245, 245);
    private static readonly Color DropValidColor = Color.FromArgb(198, 239, 206);
    private static readonly Color DropInvalidColor = Color.FromArgb(255, 199, 206);
    private static readonly Color HintIdleColor = SystemColors.GrayText;
    private static readonly Color HintValidColor = Color.FromArgb(0, 97, 0);
    private static readonly Color HintInvalidColor = Color.FromArgb(156, 0, 6);

    private readonly BindingList<ExcelRecord> _records = [];
    private readonly BindingList<ShortcutPreview> _preview = [];
    private bool _isDragActive;
    private bool _lastDropWasValid;

    public Form1()
    {
        InitializeComponent();
        gridRecords.AutoGenerateColumns = true;
        gridRecords.DataSource = _preview;
        gridRecords.DataBindingComplete += (_, _) => ApplyPreviewHeaders();
        gridRecords.RowPrePaint += gridRecords_RowPrePaint;
        EnableDragAndDrop();
        ShowDropZone();
        UpdateStatus();
    }

    private void EnableDragAndDrop()
    {
        foreach (Control control in GetDropTargets())
        {
            control.AllowDrop = true;
            control.DragEnter += DropTarget_DragEnter;
            control.DragOver += DropTarget_DragOver;
            control.DragLeave += DropTarget_DragLeave;
            control.DragDrop += DropTarget_DragDrop;
        }
    }

    private IEnumerable<Control> GetDropTargets()
    {
        yield return panelDropZone;
        yield return lblDropHint;
        yield return gridRecords;
    }

    private void ApplyPreviewHeaders()
    {
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [nameof(ShortcutPreview.Dikastirio)] = "Δικαστήριο",
            [nameof(ShortcutPreview.Afm)] = "ΑΦΜ",
            [nameof(ShortcutPreview.HostAccount)] = "Host Account",
            [nameof(ShortcutPreview.Scholia)] = "Σχόλια",
            [nameof(ShortcutPreview.Gak)] = "ΓΑΚ",
            [nameof(ShortcutPreview.Eak)] = "ΕΑΚ",
            [nameof(ShortcutPreview.Imerominia)] = "Ημερομηνία",
            [nameof(ShortcutPreview.Message)] = "Μήνυμα"
        };

        foreach (DataGridViewColumn column in gridRecords.Columns)
        {
            if (headers.TryGetValue(column.DataPropertyName, out var title))
                column.HeaderText = title;
        }
    }

    private void gridRecords_RowPrePaint(object? sender, DataGridViewRowPrePaintEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _preview.Count)
            return;

        var item = _preview[e.RowIndex];
        var row = gridRecords.Rows[e.RowIndex];
        row.DefaultCellStyle.BackColor = item.Status switch
        {
            PreviewStatus.Ready => Color.FromArgb(226, 239, 218),
            PreviewStatus.AlreadyExists => Color.FromArgb(255, 242, 204),
            PreviewStatus.Error => Color.FromArgb(252, 228, 214),
            _ => gridRecords.DefaultCellStyle.BackColor
        };
    }

    private void DropTarget_DragEnter(object? sender, DragEventArgs e) =>
        UpdateDragState(e);

    private void DropTarget_DragOver(object? sender, DragEventArgs e) =>
        UpdateDragState(e);

    private void DropTarget_DragLeave(object? sender, EventArgs e) =>
        BeginInvoke(MaybeClearDropVisual);

    private void DropTarget_DragDrop(object? sender, DragEventArgs e)
    {
        var isValid = TryGetDroppedExcelFiles(e.Data, out var files);
        ResetDropVisual();

        if (!isValid)
        {
            MessageBox.Show(
                this,
                "Επιτρέπονται μόνο αρχεία .xls ή .xlsx.",
                "Μη έγκυρο αρχείο",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        ImportFiles(files);
    }

    private void UpdateDragState(DragEventArgs e)
    {
        var isValid = TryGetDroppedExcelFiles(e.Data, out _);
        e.Effect = isValid ? DragDropEffects.Copy : DragDropEffects.None;

        if (_isDragActive && _lastDropWasValid == isValid)
            return;

        _isDragActive = true;
        _lastDropWasValid = isValid;
        ApplyDropVisual(isValid);
    }

    private void MaybeClearDropVisual()
    {
        if (IsDisposed || !IsHandleCreated)
            return;

        if (IsCursorInsideDropSurface() && (MouseButtons & MouseButtons.Left) != 0)
            return;

        ResetDropVisual();
    }

    private bool IsCursorInsideDropSurface()
    {
        var surface = panelDropZone.Visible ? (Control)panelDropZone : gridRecords;
        var clientPoint = surface.PointToClient(Cursor.Position);
        return surface.ClientRectangle.Contains(clientPoint);
    }

    private void ApplyDropVisual(bool isValid)
    {
        if (!panelDropZone.Visible)
            return;

        panelDropZone.BackColor = isValid ? DropValidColor : DropInvalidColor;
        lblDropHint.ForeColor = isValid ? HintValidColor : HintInvalidColor;
        lblDropHint.Text = isValid
            ? "Αφήστε για εισαγωγή\r\n(.xls / .xlsx)"
            : "Μη έγκυρο αρχείο\r\nΜόνο .xls ή .xlsx";
    }

    private void ResetDropVisual()
    {
        _isDragActive = false;
        _lastDropWasValid = false;

        if (!panelDropZone.Visible)
            return;

        panelDropZone.BackColor = DropIdleColor;
        lblDropHint.ForeColor = HintIdleColor;
        lblDropHint.Text = "Σύρετε αρχείο Excel εδώ\r\n(.xls ή .xlsx)";
    }

    private void ShowDropZone()
    {
        gridRecords.Visible = false;
        panelDropZone.Visible = true;
        panelDropZone.BringToFront();
        btnClear.Enabled = false;
        btnSend.Enabled = false;
        ResetDropVisual();
    }

    private void ShowPreviewGrid()
    {
        panelDropZone.Visible = false;
        gridRecords.Visible = true;
        gridRecords.BringToFront();
        btnClear.Enabled = true;
        btnSend.Enabled = _preview.Any(p => p.CanSend);
        _isDragActive = false;
        _lastDropWasValid = false;
    }

    private static bool TryGetDroppedExcelFiles(IDataObject? data, out string[] files)
    {
        files = [];
        if (data is null || !data.GetDataPresent(DataFormats.FileDrop))
            return false;

        if (data.GetData(DataFormats.FileDrop) is not string[] dropped || dropped.Length == 0)
            return false;

        var excelFiles = dropped
            .Where(path => AllowedExtensions.Contains(Path.GetExtension(path)))
            .ToArray();

        if (excelFiles.Length == 0 || excelFiles.Length != dropped.Length)
            return false;

        files = excelFiles;
        return true;
    }

    private void btnAddExcel_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Επιλογή Excel",
            Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx",
            Multiselect = true
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        ImportFiles(dialog.FileNames);
    }

    private void ImportFiles(IEnumerable<string> paths)
    {
        var added = 0;
        var errors = new List<string>();

        foreach (var path in paths)
        {
            var extension = Path.GetExtension(path);
            if (!AllowedExtensions.Contains(extension))
            {
                errors.Add($"{Path.GetFileName(path)}: Επιτρέπονται μόνο .xls ή .xlsx.");
                continue;
            }

            try
            {
                var rows = ExcelImportService.Import(path);
                foreach (var row in rows)
                    _records.Add(row);

                added += rows.Count;
            }
            catch (MissingExcelFieldsException ex)
            {
                errors.Add($"{Path.GetFileName(path)}: Δεν βρέθηκαν όλα τα πεδία ({string.Join(", ", ex.MissingFields)}).");
            }
            catch (Exception ex)
            {
                errors.Add($"{Path.GetFileName(path)}: {ex.Message}");
            }
        }

        if (errors.Count > 0)
        {
            MessageBox.Show(
                this,
                string.Join(Environment.NewLine, errors),
                "Σφάλμα εισαγωγής",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        else if (added == 0)
        {
            MessageBox.Show(
                this,
                "Δεν βρέθηκαν γραμμές δεδομένων στο επιλεγμένο αρχείο.",
                "Εισαγωγή",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        if (_records.Count > 0)
            BuildAndShowPreview();
        else
            ShowDropZone();

        UpdateStatus();
    }

    private void BuildAndShowPreview()
    {
        AppSettings settings;
        try
        {
            settings = AppSettings.Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"Αδυναμία φόρτωσης ρυθμίσεων (appsettings.json): {ex.Message}",
                "Ρυθμίσεις",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        if (!settings.CanRunFolderLookup)
        {
            MessageBox.Show(
                this,
                "Ορίστε το ConnectionStrings:CTCOLLECT στο appsettings.json.",
                "Σύνδεση βάσης",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        try
        {
            Cursor = Cursors.WaitCursor;
            var fileService = new ImpersonatedFileService(settings);
            var lookup = new FolderLookupService(settings, fileService);
            var shortcutService = new ShortcutService(lookup, fileService);
            var previewRows = shortcutService.BuildPreview(_records);

            _preview.Clear();
            foreach (var row in previewRows)
                _preview.Add(row);

            ShowPreviewGrid();

            var errorRows = _preview.Where(p => p.Status == PreviewStatus.Error).ToList();
            if (errorRows.Count > 0 && !_preview.Any(p => p.CanSend))
            {
                var details = string.Join(
                    Environment.NewLine,
                    errorRows.Take(5).Select(e => $"• {e.HostAccount} / {e.Gak}: {e.Message}"));
                MessageBox.Show(
                    this,
                    "Δεν υπάρχουν έτοιμες εγγραφές για αποστολή." + Environment.NewLine + Environment.NewLine + details,
                    "Preview — σφάλματα",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"Σφάλμα preview: {ex.Message}",
                "Preview",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
            UpdateStatus();
        }
    }

    private void btnSend_Click(object? sender, EventArgs e)
    {
        var ready = _preview.Where(p => p.CanSend).ToList();
        if (ready.Count == 0)
        {
            MessageBox.Show(
                this,
                "Δεν υπάρχουν έτοιμες συντομεύσεις για αποστολή.",
                "Αποστολή",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        var confirm = MessageBox.Show(
            this,
            $"Να σταλούν {ready.Count} συντομεύσεις στους φακέλους;",
            "Επιβεβαίωση αποστολής",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
            return;

        AppSettings settings;
        try
        {
            settings = AppSettings.Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Ρυθμίσεις", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            Cursor = Cursors.WaitCursor;
            var fileService = new ImpersonatedFileService(settings);
            var lookup = new FolderLookupService(settings, fileService);
            var shortcutService = new ShortcutService(lookup, fileService);
            var results = shortcutService.Send(ready);

            var created = results.Count(r => r.Action == ShortcutAction.Created);
            var skipped = results.Count(r => r.Action == ShortcutAction.SkippedExisting);
            var failed = results.Where(r => r.Action == ShortcutAction.Failed).ToList();

            // Ανανέωση preview μετά την αποστολή
            BuildAndShowPreview();

            var summary = $"Αποστολή ολοκληρώθηκε: {created} νέες, {skipped} υπήρχαν ήδη, {failed.Count} αποτυχίες.";
            MessageBox.Show(
                this,
                failed.Count == 0
                    ? summary
                    : summary + Environment.NewLine + Environment.NewLine +
                      string.Join(Environment.NewLine, failed.Take(10).Select(f => $"• {f.HostAccount}: {f.Message}")),
                "Αποστολή",
                MessageBoxButtons.OK,
                failed.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Αποστολή", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void btnClear_Click(object? sender, EventArgs e)
    {
        if (_records.Count == 0 && _preview.Count == 0)
            return;

        var result = MessageBox.Show(
            this,
            "Να καθαριστεί το preview;",
            "Επιβεβαίωση",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
            return;

        _records.Clear();
        _preview.Clear();
        UpdateStatus();
        ShowDropZone();
    }

    private void UpdateStatus()
    {
        var ready = _preview.Count(p => p.CanSend);
        var errors = _preview.Count(p => p.Status == PreviewStatus.Error);
        var existing = _preview.Count(p => p.Status == PreviewStatus.AlreadyExists);

        lblStatus.Text = _preview.Count == 0
            ? $"Εγγραφές: {_records.Count}"
            : $"Preview: {_preview.Count}  |  Έτοιμα: {ready}  |  Υπάρχουν: {existing}  |  Σφάλματα: {errors}";

        btnSend.Enabled = ready > 0;
    }
}
