using ExcelImporter.Models;
using Microsoft.Data.SqlClient;

namespace ExcelImporter.Services;

public sealed class FolderMatch
{
    public string Fakelos { get; init; } = string.Empty;
    public string NewFakelos2 { get; init; } = string.Empty;
    public string TargetDirectory { get; init; } = string.Empty;
}

public sealed class FolderLookupService
{
    private readonly string _connectionString;
    private readonly string _scannedRoot;
    private readonly ImpersonatedFileService _fileService;

    public FolderLookupService(AppSettings settings, ImpersonatedFileService fileService)
    {
        _connectionString = settings.ConnectionString;
        _scannedRoot = settings.ScannedRoot;
        _fileService = fileService;
    }

    public FolderMatch Resolve(ExcelRecord record)
    {
        if (string.IsNullOrWhiteSpace(record.HostAccount))
            throw new InvalidOperationException("Λείπει το Host Account.");

        if (string.IsNullOrWhiteSpace(record.Gak))
            throw new InvalidOperationException("Λείπει το ΓΑΚ.");

        if (string.IsNullOrWhiteSpace(record.Onomateponymo))
            throw new InvalidOperationException("Λείπει το Ονοματεπώνυμο.");

        var namePrefix = GetNamePrefix(record.Onomateponymo);

        using var connection = new SqlConnection(_connectionString);
        connection.Open();

        var contractNullable = GetContractFolders(connection, record.HostAccount);
        if (contractNullable is null)
        {
            throw new InvalidOperationException(
                $"Δεν βρέθηκε CONTRACTS_PLUS για contractnbr={record.HostAccount}.");
        }

        var contract = contractNullable.Value;

        var dikastiriaFakelos = GetDikastiriaFakelos(
            connection,
            record.Gak,
            namePrefix,
            record.HostAccount);

        if (dikastiriaFakelos is null)
        {
            throw new InvalidOperationException(
                $"Δεν βρέθηκε Dikastiria για GAK={record.Gak}, όνομα={namePrefix}%, LiveContractNbr={record.HostAccount}.");
        }

        if (!string.Equals(dikastiriaFakelos, contract.NewFakelos2, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Μη έγκυρος φάκελος: Dikastiria.fakelos='{dikastiriaFakelos}' ≠ CONTRACTS_PLUS.newfakelos2='{contract.NewFakelos2}'.");
        }

        if (string.IsNullOrWhiteSpace(contract.Fakelos) || string.IsNullOrWhiteSpace(contract.NewFakelos2))
            throw new InvalidOperationException("Κενές τιμές fakelos/newfakelos2 από CONTRACTS_PLUS.");

        var target = Path.Combine(_scannedRoot, contract.NewFakelos2, contract.Fakelos);
        if (!_fileService.DirectoryExistsWithImpersonate(target, out var dirError))
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(dirError)
                ? $"Ο φάκελος προορισμού δεν υπάρχει: {target}"
                : dirError);
        }

        return new FolderMatch
        {
            Fakelos = contract.Fakelos,
            NewFakelos2 = contract.NewFakelos2,
            TargetDirectory = target
        };
    }

    private static (string Fakelos, string NewFakelos2)? GetContractFolders(SqlConnection connection, string contractNbr)
    {
        const string sql = """
            SELECT TOP 1 fakelos, newfakelos2
            FROM CONTRACTS_PLUS
            WHERE contractnbr = @contractnbr
            """;

        using var cmd = new SqlCommand(sql, connection);
        cmd.Parameters.AddWithValue("@contractnbr", contractNbr);

        using var reader = cmd.ExecuteReader();
        if (!reader.Read())
            return null;

        var fakelos = reader["fakelos"]?.ToString()?.Trim() ?? string.Empty;
        var newFakelos2 = reader["newfakelos2"]?.ToString()?.Trim() ?? string.Empty;
        return (fakelos, newFakelos2);
    }

    private static string? GetDikastiriaFakelos(
        SqlConnection connection,
        string gak,
        string namePrefix,
        string liveContractNbr)
    {
        const string sql = """
            SELECT TOP 1 fakelos, fullname, LiveContractNbr
            FROM Dikastiria
            WHERE GAK = @gak
              AND FullName LIKE @namePrefix
              AND LiveContractNbr = @liveContractNbr
            """;

        using var cmd = new SqlCommand(sql, connection);
        cmd.Parameters.AddWithValue("@gak", gak);
        cmd.Parameters.AddWithValue("@namePrefix", namePrefix + "%");
        cmd.Parameters.AddWithValue("@liveContractNbr", liveContractNbr);

        using var reader = cmd.ExecuteReader();
        if (!reader.Read())
            return null;

        return reader["fakelos"]?.ToString()?.Trim();
    }

    public static string GetNamePrefix(string fullName)
    {
        var cleaned = fullName.Trim();
        if (cleaned.Length == 0)
            return string.Empty;

        return cleaned.Length <= 5 ? cleaned : cleaned[..5];
    }
}
