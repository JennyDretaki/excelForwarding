using Microsoft.Extensions.Configuration;

namespace ExcelImporter.Services;

public sealed class AppSettings
{
    public string ConnectionString { get; }
    public string ScannedRoot { get; }
    public string ImpDomain { get; }
    public string ImpUser { get; }
    public string ImpPassword { get; }

    private AppSettings(
        string connectionString,
        string scannedRoot,
        string impDomain,
        string impUser,
        string impPassword)
    {
        ConnectionString = connectionString;
        ScannedRoot = scannedRoot;
        ImpDomain = impDomain;
        ImpUser = impUser;
        ImpPassword = impPassword;
    }

    public static AppSettings Load()
    {
        var basePath = AppContext.BaseDirectory;
        var config = new ConfigurationBuilder()
            .SetBasePath(basePath)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .Build();

        // Προτεραιότητα: CTCOLLECT, αλλιώς ExpensesDb (παλιό όνομα)
        var connectionString = config.GetConnectionString("CTCOLLECT")?.Trim()
            ?? config.GetConnectionString("ExpensesDb")?.Trim()
            ?? string.Empty;

        var scannedRoot = config["Paths:ScannedRoot"]?.Trim();
        if (string.IsNullOrWhiteSpace(scannedRoot))
            scannedRoot = @"\\srv01\public\expenses\SCANNED";

        scannedRoot = ImpersonatedFileService.ToUncPath(scannedRoot);

        return new AppSettings(
            connectionString,
            scannedRoot,
            config["Impersonation:Domain"]?.Trim() ?? string.Empty,
            config["Impersonation:User"]?.Trim() ?? string.Empty,
            config["Impersonation:Password"] ?? string.Empty);
    }

    public bool HasValidConnectionString =>
        !string.IsNullOrWhiteSpace(ConnectionString)
        && !ConnectionString.Contains("YOUR_SERVER", StringComparison.OrdinalIgnoreCase)
        && !ConnectionString.Contains("YOUR_DATABASE", StringComparison.OrdinalIgnoreCase);

    public bool HasImpersonationCredentials =>
        !string.IsNullOrWhiteSpace(ImpDomain)
        && !string.IsNullOrWhiteSpace(ImpUser)
        && !string.IsNullOrWhiteSpace(ImpPassword);

    public bool CanRunFolderLookup => HasValidConnectionString;
}
