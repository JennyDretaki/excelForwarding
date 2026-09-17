using System.Runtime.InteropServices;
using System.Security.Principal;

namespace ExcelImporter.Services;

/// <summary>
/// Αντιγραφή αρχείων με Windows impersonation (CBS Application_Tools style).
/// Χρησιμοποιεί NEW_CREDENTIALS για πρόσβαση σε network/UNC paths.
/// </summary>
public sealed class ImpersonatedFileService
{
    // καλύτερο για network shares
    private const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
    private const int LOGON32_PROVIDER_WINNT50 = 3;

    private readonly string _domain;
    private readonly string _user;
    private readonly string _password;

    public ImpersonatedFileService(AppSettings settings)
    {
        if (!settings.HasImpersonationCredentials)
            throw new InvalidOperationException("Λείπουν τα credentials impersonation από το appsettings.json.");

        _domain = settings.ImpDomain;
        _user = settings.ImpUser;
        _password = settings.ImpPassword;
    }

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool LogonUser(
        string lpszUsername,
        string lpszDomain,
        string lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        out IntPtr phToken);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    public static MemoryStream FileToMemoryStream(string filePath)
    {
        var bytes = File.ReadAllBytes(filePath);
        var ms = new MemoryStream(bytes);
        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Μετατρέπει W:\... σε UNC \\srv01\public\... (τα mapped drives δεν φαίνονται με impersonation).
    /// </summary>
    public static string ToUncPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return path;

        if (path.StartsWith(@"W:\", StringComparison.OrdinalIgnoreCase))
            return @"\\srv01\public\" + path[3..];

        if (path.StartsWith(@"W:/", StringComparison.OrdinalIgnoreCase))
            return @"\\srv01\public\" + path[3..].Replace('/', '\\');

        return path;
    }

    public static string ApplyLongPathPrefix(string filedst)
    {
        filedst = ToUncPath(filedst);

        if (filedst.Length <= 200)
            return filedst;

        if (filedst.StartsWith(@"\\", StringComparison.Ordinal))
        {
            // \\server\share\... → \\?\UNC\server\share\...
            return @"\\?\UNC\" + filedst.TrimStart('\\');
        }

        return @"\\?\" + filedst;
    }

    public bool CopyFileFromStreamWithImpersonate(string destFolder, MemoryStream ms, string fileDst, out string error)
    {
        destFolder = ToUncPath(destFolder);
        fileDst = ToUncPath(fileDst);

        error = string.Empty;
        var localError = string.Empty;
        var ret = false;

        if (!TryImpersonate(out var tokenHandle, out error))
            return false;

        try
        {
            using var newId = new WindowsIdentity(tokenHandle);
            WindowsIdentity.RunImpersonated(newId.AccessToken, () =>
            {
                try
                {
                    if (!Directory.Exists(destFolder))
                    {
                        localError = $"Ο φάκελος προορισμού δεν υπάρχει: {destFolder}";
                        return;
                    }

                    File.WriteAllBytes(ApplyLongPathPrefix(fileDst), ms.ToArray());
                    ret = true;
                }
                catch (Exception ex)
                {
                    localError = "File access error: " + ex.Message;
                }
            });
        }
        finally
        {
            CloseHandle(tokenHandle);
        }

        error = localError;
        return ret;
    }

    public bool FileExistsWithImpersonate(string path, out string error)
    {
        path = ToUncPath(path);
        error = string.Empty;
        var localError = string.Empty;
        var exists = false;

        if (!TryImpersonate(out var tokenHandle, out error))
            return false;

        try
        {
            using var newId = new WindowsIdentity(tokenHandle);
            WindowsIdentity.RunImpersonated(newId.AccessToken, () =>
            {
                try
                {
                    exists = File.Exists(ApplyLongPathPrefix(path));
                }
                catch (Exception ex)
                {
                    localError = ex.Message;
                }
            });
        }
        finally
        {
            CloseHandle(tokenHandle);
        }

        error = localError;
        return exists;
    }

    public bool DirectoryExistsWithImpersonate(string path, out string error)
    {
        path = ToUncPath(path);
        error = string.Empty;
        var localError = string.Empty;
        var exists = false;

        if (!TryImpersonate(out var tokenHandle, out error))
            return false;

        try
        {
            using var newId = new WindowsIdentity(tokenHandle);
            WindowsIdentity.RunImpersonated(newId.AccessToken, () =>
            {
                try
                {
                    exists = Directory.Exists(path);
                }
                catch (Exception ex)
                {
                    localError = ex.Message;
                }
            });
        }
        finally
        {
            CloseHandle(tokenHandle);
        }

        error = localError;
        return exists;
    }

    private bool TryImpersonate(out IntPtr tokenHandle, out string error)
    {
        error = string.Empty;
        var ok = LogonUser(
            _user,
            _domain,
            _password,
            LOGON32_LOGON_NEW_CREDENTIALS,
            LOGON32_PROVIDER_WINNT50,
            out tokenHandle);

        if (ok)
            return true;

        error = $"LogonUser failed. Error code: {Marshal.GetLastWin32Error()}";
        return false;
    }
}
