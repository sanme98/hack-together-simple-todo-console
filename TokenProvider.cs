// Licensed under the MIT License.

using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace GraphTutorial;

public class TokenProvider : IAccessTokenProvider
{
    private readonly IPublicClientApplication publicClientApplication;
    public TokenProvider(string clientId, string tenantId)
    {
        publicClientApplication = PublicClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantId)
            .Build();
        AllowedHostsValidator = new AllowedHostsValidator();

        // Building StorageCreationProperties
        var storageProperties =
             new StorageCreationPropertiesBuilder(CacheSettings.CacheFileName, CacheSettings.CacheDir)
             .WithLinuxKeyring(
                 CacheSettings.LinuxKeyRingSchema,
                 CacheSettings.LinuxKeyRingCollection,
                 CacheSettings.LinuxKeyRingLabel,
                 CacheSettings.LinuxKeyRingAttr1,
                 CacheSettings.LinuxKeyRingAttr2)
             .WithMacKeyChain(
                 CacheSettings.KeyChainServiceName,
                 CacheSettings.KeyChainAccountName)
             .Build();

        // This hooks up the cross-platform cache into MSAL
        var cacheHelper = CreateCacheHelper(storageProperties).Result; //await MsalCacheHelper.CreateAsync(storageProperties);
        cacheHelper.RegisterCache(publicClientApplication.UserTokenCache);
    }

    public async Task ClearCacheAsync()
    {
        var accounts3 = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);
        foreach (var acc in accounts3)
        {
            Console.WriteLine($"Removing account for {acc.Username}");
            Console.Write(Environment.NewLine);
            await publicClientApplication.RemoveAsync(acc).ConfigureAwait(false);
        }
    }

    public static async Task<MsalCacheHelper> CreateCacheHelper(StorageCreationProperties storageProperties)
    {
        return await MsalCacheHelper.CreateAsync(storageProperties); ;
    }

    public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
        CancellationToken cancellationToken = default)
    {
        var scopes = new[] { "User.Read" };

        AuthenticationResult result;
        try
        {
            var accounts = await publicClientApplication.GetAccountsAsync();

            // Try to acquire an access token from the cache. If an interaction is required, 
            // MsalUiRequiredException will be thrown.
            result = await publicClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault()).ExecuteAsync();
            return await Task.FromResult(result.AccessToken);
        }
        catch (MsalUiRequiredException)
        {
            // Acquiring an access token interactively. MSAL will cache it so we can use AcquireTokenSilent
            // on future calls.
            var result1 = await publicClientApplication.AcquireTokenWithDeviceCode(scopes, deviceCodeResult =>
            {
                // This will print the message on the console which tells the user where to go sign-in using
                // a separate browser and the code to enter once they sign in.
                // The AcquireTokenWithDeviceCode() method will poll the server after firing this
                // device code callback to look for the successful login of the user via that browser.
                // This background polling (whose interval and timeout data is also provided as fields in the
                // deviceCodeCallback class) will occur until:
                // * The user has successfully logged in via browser and entered the proper code
                // * The timeout specified by the server for the lifetime of this code (typically ~15 minutes) has been reached
                // * The developing application calls the Cancel() method on a CancellationToken sent into the method.
                //   If this occurs, an OperationCanceledException will be thrown (see catch below for more details).
                Console.WriteLine(deviceCodeResult.Message);
                return Task.FromResult(0);
            }).ExecuteAsync();
            // get the token and return it in your own way
            return await Task.FromResult(result1.AccessToken);
        }
    }

    public AllowedHostsValidator AllowedHostsValidator { get; }
}
