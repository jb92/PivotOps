using Windows.Security.Credentials;

namespace PivotOps.Desktop.Services;

/// <summary>
/// Backs <c>OfficeRuntime.storage</c> with the Windows Credential Locker so the
/// Azure DevOps PAT never lands in WebView2's localStorage.
/// </summary>
internal sealed class SecureStore
{
    private const string Resource = "PivotOps";
    private readonly PasswordVault _vault = new();

    public string? Get(string key)
    {
        try
        {
            var credential = _vault.Retrieve(Resource, key);
            credential.RetrievePassword();
            return credential.Password;
        }
        catch
        {
            return null;
        }
    }

    public void Set(string key, string value)
    {
        Remove(key);
        _vault.Add(new PasswordCredential(Resource, key, value));
    }

    public void Remove(string key)
    {
        try
        {
            _vault.Remove(_vault.Retrieve(Resource, key));
        }
        catch
        {
            // Nothing stored under this key.
        }
    }
}
