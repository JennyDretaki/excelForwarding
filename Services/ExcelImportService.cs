using System.Data;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using ExcelDataReader;
using ExcelImporter.Models;

namespace ExcelImporter.Services;

public sealed class MissingExcelFieldsException : Exception
{
    public IReadOnlyList<string> MissingFields { get; }

    public MissingExcelFieldsException(IEnumerable<string> missingFields)
        : base($"Δεν βρέθηκαν όλα τα πεδία ({string.Join(", ", missingFields)}).")
    {
        MissingFields = missingFields.ToArray();
    }
}

public static class ExcelImportService
{
    private static readonly (string DisplayName, string[] Candidates)[] RequiredFields =
    [
        ("Δικαστήριο", ["ΔΙΚΑΣΤΗΡΙΟ"]),
        ("Ονοματεπώνυμο", ["ΟΝΟΜΑΤΕΠΩΝΥΜΟ"]),
        ("ΑΦΜ", ["ΑΦΜ"]),
        ("Host Account", ["Host Account", "HOST ACCOUNT"]),
        ("Σχόλια", ["ΣΧΟΛΙΟ/ΕΝΕΡΓΕΙΑ", "ΣΧΟΛΙΑ CEPAL", "ΣΧΟΛΙΑ"]),
        ("ΓΑΚ", ["ΓΑΚ"]),
        ("ΕΑΚ", ["ΕΑΚ"]),
        ("Ημερομηνία", ["ΗΜΕΡΟΜΗΝΙΑ ΔΙΚΑΣΙΜΟΥ", "ΗΜΕΡΟΜΗΝΙΑ"])
    ];

    static ExcelImportService()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static IReadOnlyList<ExcelRecord> Import(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            return ImportWithClosedXml(filePath);

        if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            return ImportWithExcelDataReader(filePath);

        throw new InvalidOperationException("Επιτρέπονται μόνο αρχεία .xls ή .xlsx.");
    }

    private static IReadOnlyList<ExcelRecord> ImportWithClosedXml(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();
        var used = worksheet.RangeUsed();
        if (used is null)
            return [];

        var map = BuildHeaderMap(used.FirstRow());
        EnsureRequiredFields(map);
        var columns = ResolveRequiredColumns(map);
        var records = new List<ExcelRecord>();

        foreach (var row in used.RowsUsed().Skip(1))
        {
            var record = CreateRecord(
                GetText(row, columns["Δικαστήριο"]),
                GetText(row, columns["Ονοματεπώνυμο"]),
                GetText(row, columns["ΑΦΜ"]),
                GetText(row, columns["Host Account"]),
                GetText(row, columns["Σχόλια"]),
                GetText(row, columns["ΓΑΚ"]),
                GetText(row, columns["ΕΑΚ"]),
                GetDateText(row, columns["Ημερομηνία"]),
                filePath);

            if (!IsEmpty(record))
                records.Add(record);
        }

        return records;
    }

    private static IReadOnlyList<ExcelRecord> ImportWithExcelDataReader(string filePath)
    {
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true
            }
        });

        if (dataSet.Tables.Count == 0)
            return [];

        var table = dataSet.Tables[0];
        var map = BuildHeaderMap(table);
        EnsureRequiredFields(map);
        var columns = ResolveRequiredColumns(map);
        var records = new List<ExcelRecord>();

        foreach (DataRow row in table.Rows)
        {
            var record = CreateRecord(
                GetText(row, columns["Δικαστήριο"]),
                GetText(row, columns["Ονοματεπώνυμο"]),
                GetText(row, columns["ΑΦΜ"]),
                GetText(row, columns["Host Account"]),
                GetText(row, columns["Σχόλια"]),
                GetText(row, columns["ΓΑΚ"]),
                GetText(row, columns["ΕΑΚ"]),
                GetDateText(row, columns["Ημερομηνία"]),
                filePath);

            if (!IsEmpty(record))
                records.Add(record);
        }

        return records;
    }

    private static ExcelRecord CreateRecord(
        string dikastirio,
        string onomateponymo,
        string afm,
        string hostAccount,
        string scholia,
        string gak,
        string eak,
        string imerominia,
        string filePath) =>
        new()
        {
            Dikastirio = dikastirio,
            Onomateponymo = onomateponymo,
            Afm = afm,
            HostAccount = hostAccount,
            Scholia = scholia,
            Gak = gak,
            Eak = eak,
            Imerominia = imerominia,
            SourceFilePath = filePath
        };

    private static Dictionary<string, int> BuildHeaderMap(IXLRangeRow headerRow)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in headerRow.CellsUsed())
        {
            var header = Normalize(cell.GetString());
            if (string.IsNullOrWhiteSpace(header))
                continue;

            map[header] = cell.Address.ColumnNumber;
        }

        return map;
    }

    private static Dictionary<string, int> BuildHeaderMap(DataTable table)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < table.Columns.Count; i++)
        {
            var header = Normalize(table.Columns[i].ColumnName);
            if (string.IsNullOrWhiteSpace(header))
                continue;

            map[header] = i;
        }

        return map;
    }

    private static void EnsureRequiredFields(Dictionary<string, int> map)
    {
        var missing = RequiredFields
            .Where(field => !field.Candidates.Any(candidate => map.ContainsKey(Normalize(candidate))))
            .Select(field => field.DisplayName)
            .ToList();

        if (missing.Count > 0)
            throw new MissingExcelFieldsException(missing);
    }

    private static Dictionary<string, int> ResolveRequiredColumns(Dictionary<string, int> map)
    {
        var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var field in RequiredFields)
        {
            foreach (var candidate in field.Candidates)
            {
                if (map.TryGetValue(Normalize(candidate), out var col))
                {
                    columns[field.DisplayName] = col;
                    break;
                }
            }
        }

        return columns;
    }

    private static string Normalize(string value) =>
        value.Trim().Replace('\u00A0', ' ');

    private static string GetText(IXLRangeRow row, int column)
    {
        var cell = row.Cell(column);
        if (cell.IsEmpty())
            return string.Empty;

        return cell.GetFormattedString().Trim();
    }

    private static string GetText(DataRow row, int column)
    {
        if (column < 0 || column >= row.Table.Columns.Count)
            return string.Empty;

        var value = row[column];
        if (value is null or DBNull)
            return string.Empty;

        return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
    }

    private static string GetDateText(IXLRangeRow row, int column)
    {
        var cell = row.Cell(column);
        if (cell.IsEmpty())
            return string.Empty;

        if (cell.TryGetValue(out DateTime date))
            return date.ToString("dd/MM/yyyy");

        if (cell.DataType == XLDataType.Number && cell.TryGetValue(out double serial))
        {
            try
            {
                return DateTime.FromOADate(serial).ToString("dd/MM/yyyy");
            }
            catch
            {
                // fall through
            }
        }

        return cell.GetFormattedString().Trim();
    }

    private static string GetDateText(DataRow row, int column)
    {
        if (column < 0 || column >= row.Table.Columns.Count)
            return string.Empty;

        var value = row[column];
        if (value is null or DBNull)
            return string.Empty;

        if (value is DateTime date)
            return date.ToString("dd/MM/yyyy");

        if (value is double serial)
        {
            try
            {
                return DateTime.FromOADate(serial).ToString("dd/MM/yyyy");
            }
            catch
            {
                // fall through
            }
        }

        if (DateTime.TryParse(Convert.ToString(value, CultureInfo.CurrentCulture), out var parsed))
            return parsed.ToString("dd/MM/yyyy");

        return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
    }

    private static bool IsEmpty(ExcelRecord record) =>
        string.IsNullOrWhiteSpace(record.Dikastirio)
        && string.IsNullOrWhiteSpace(record.Onomateponymo)
        && string.IsNullOrWhiteSpace(record.Afm)
        && string.IsNullOrWhiteSpace(record.HostAccount)
        && string.IsNullOrWhiteSpace(record.Scholia)
        && string.IsNullOrWhiteSpace(record.Gak)
        && string.IsNullOrWhiteSpace(record.Eak)
        && string.IsNullOrWhiteSpace(record.Imerominia);
}
