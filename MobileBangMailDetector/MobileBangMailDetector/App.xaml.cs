using System;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MobileBangMailDetector
{
    public partial class App : Xamarin.Forms.Application
    {
        // UIParent used by Android version of the app
        public static object AuthUIParent = null;

        // Keychain security group used by iOS version of the app
        public static string iOSKeychainSecurityGroup = null;

        // Microsoft Authentication client for native/mobile apps
        public static IPublicClientApplication PCA;

        // Microsoft Graph client
        public static GraphServiceClient GraphClient;

        // Microsoft Graph permissions used by app
        private readonly string[] Scopes = OAuthSettings.Scopes.Split(' ');

        public App()
        {
            InitializeComponent();

            var builder = PublicClientApplicationBuilder
               .Create(OAuthSettings.ApplicationId)
               .WithRedirectUri(OAuthSettings.RedirectUri);

            if (!string.IsNullOrEmpty(iOSKeychainSecurityGroup))
            {
                builder = builder.WithIosKeychainSecurityGroup(iOSKeychainSecurityGroup);
            }

            PCA = builder.Build();

            MainPage = new MainPage();
        }

        public async Task SignIn()
        {
            // First, attempt silent sign in
            // If the user's information is already in the app's cache,
            // they won't have to sign in again.
            try
            {
                var accounts = await PCA.GetAccountsAsync();

                var silentAuthResult = await PCA
                    .AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();

                Debug.WriteLine("User already signed in.");
                Debug.WriteLine($"Successful silent authentication for: {silentAuthResult.Account.Username}");
                Debug.WriteLine($"Access token: {silentAuthResult.AccessToken}");
            }
            catch (MsalUiRequiredException msalEx)
            {
                // This exception is thrown when an interactive sign-in is required.
                Debug.WriteLine("Silent token request failed, user needs to sign-in: " + msalEx.Message);
                // Prompt the user to sign-in
                var interactiveRequest = PCA.AcquireTokenInteractive(Scopes);

                if (AuthUIParent != null)
                {
                    interactiveRequest = interactiveRequest
                        .WithParentActivityOrWindow(AuthUIParent);
                }

                var interactiveAuthResult = await interactiveRequest.ExecuteAsync();
                Debug.WriteLine($"Successful interactive authentication for: {interactiveAuthResult.Account.Username}");
                Debug.WriteLine($"Access token: {interactiveAuthResult.AccessToken}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Authentication failed. See exception messsage for more details: " + ex.Message);
            }

            await InitializeGraphClientAsync();
        }

        public async Task InitializeGraphClientAsync()
        {
            var currentAccounts = await PCA.GetAccountsAsync();
            try
            {
                if (currentAccounts.Count() > 0)
                {
                    // Initialize Graph client
                    GraphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            var result = await PCA.AcquireTokenSilent(Scopes, currentAccounts.FirstOrDefault())
                                .ExecuteAsync();

                            requestMessage.Headers.Authorization =
                                new AuthenticationHeaderValue("Bearer", result.AccessToken);
                        }));

                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to initialized graph client.");
                Debug.WriteLine($"Accounts in the msal cache: {currentAccounts.Count()}.");
                Debug.WriteLine($"See exception message for details: {ex.Message}");
            }
        }

        protected override void OnSleep()
        {
        }

        protected override void OnResume()
        {
        }
    }
}
