using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Configuration;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace MSGraphMailer.GraphHelper
{
    /// <summary>
    /// Instantiates a new GraphServiceClient and Token.
    /// </summary>
    internal class GetGraphServiceClient
    {
        /* 
         * Example of AppSecret.config
         *  <appSettings>
	     *      <!-- Azure App MS Graph config: -->
	     *      <add key = "tenantId" value = "00000000-0000-0000-0000-000000000000" />
	     *      <add key = "clientId" value = "00000000-0000-0000-0000-000000000000" />
	     *      <!-- byClientSecret: -->
	     *      <add key = "clientSecret" value = "--qwertyuiopasdfghjklzxcvbnm0123456789--" />
	     *      <!-- byCertificatePath: -->
	     *      <add key = "ertificatePath" value = "*.pfx" />
	     *      <add key = "certificatePass" value = "password" />
	     *      <!-- byCertificateThumbprint: -->
	     *      <add key = "certificateThumbp1rint" value = "QWERTYUIOPASDFGHJKLZXCVBNM0123456789QWER" />
         *  </appSettings>
        */

        private string _tenantId;
        private string _clientId;
        private string _clientSecret;

        private string _certificatePath;
        private string _certificatePass;

        private string _certificateThumbprint;

        /// <summary>
        /// Returns a Graph service client.
        /// </summary>
        public GraphServiceClient GraphClient { get; private set; }
        /// <summary>
        /// Returns a Token for Azure App.
        /// </summary>
        public string Token { get; private set; }
        
        /// <summary>
        /// GetGraphServiceClient exception.
        /// </summary>
        public Exception Exception { get; private set; }

        /// <summary>
        /// Instantiates a new GraphServiceClient and Token.
        /// </summary>
        /// <param name="authorizedBy">Enum how to authorized in Azure.</param>
        public GetGraphServiceClient(AuthorizedBy authorizedBy = AuthorizedBy.byClientSecret)
        {
            if (GetMSGraphConfig(authorizedBy))
            {
                Token = GetAccessToken(authorizedBy).Result;
                GraphClient = GetGraphClient();
            }
        }

        /// <summary>
        /// Read config file and prepare local files for authorization.
        /// </summary>
        /// <param name="authorizedBy">Enum how to authorized in Azure.</param>
        /// <returns></returns>
        private Boolean GetMSGraphConfig(AuthorizedBy authorizedBy)
        {
            var getMSGraphConfig = false;
            try
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains("tenantId") && ConfigurationManager.AppSettings.AllKeys.Contains("clientId"))
                {
                    _tenantId = ConfigurationManager.AppSettings.Get("tenantId");
                    _clientId = ConfigurationManager.AppSettings.Get("clientId");
                    if (authorizedBy == AuthorizedBy.byClientSecret && ConfigurationManager.AppSettings.AllKeys.Contains("clientSecret"))
                    {
                        _clientSecret = ConfigurationManager.AppSettings.Get("clientSecret");
                    }
                    else if (authorizedBy == AuthorizedBy.byCertificatePath && ConfigurationManager.AppSettings.AllKeys.Contains("certificatePath") && ConfigurationManager.AppSettings.AllKeys.Contains("certificatePass"))
                    {
                        _certificatePath = ConfigurationManager.AppSettings.Get("certificatePath");
                        _certificatePass = ConfigurationManager.AppSettings.Get("certificatePass");
                    }
                    else if (authorizedBy == AuthorizedBy.byCertificateThumbprint && ConfigurationManager.AppSettings.AllKeys.Contains("certificateThumbprint"))
                    {
                        _certificateThumbprint = ConfigurationManager.AppSettings.Get("certificateThumbprint");
                    }
                    else
                    {
                        throw new Exception($"Failed to read config for {authorizedBy}.");
                    }
                    getMSGraphConfig = true;
                }
            }
            catch (Exception ex)
            {
                Exception = ex;
            }
            return getMSGraphConfig;
        }

        /// <summary>
        /// Task for authorization in Azure.
        /// </summary>
        /// <param name="authorizedBy">Enum how to authorized in Azure.</param>
        /// <returns></returns>
        private async Task<string> GetAccessToken(AuthorizedBy authorizedBy)
        {
            try
            {
                IConfidentialClientApplication confidentialClient = null;
                if (authorizedBy == AuthorizedBy.byClientSecret)
                {
                    confidentialClient = ConfidentialClientApplicationBuilder
                        .Create(_clientId)
                        .WithClientSecret(_clientSecret)
                        .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}/v2.0"))
                        .Build();
                }
                else if (authorizedBy == AuthorizedBy.byCertificatePath)
                {
                    confidentialClient = ConfidentialClientApplicationBuilder
                        .Create(_clientId)
                        .WithCertificate(GetCertificateFromDirectory(_certificatePath, _certificatePass))
                        .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}/v2.0"))
                        .Build();
                }
                else if (authorizedBy == AuthorizedBy.byCertificateThumbprint)
                {
                    confidentialClient = ConfidentialClientApplicationBuilder
                        .Create(_clientId)
                        .WithCertificate(GetCertificateFromStore(_certificateThumbprint))
                        .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}/v2.0"))
                        .Build();
                }

                if (confidentialClient != null)
                {
                    var authResult = await confidentialClient
                            .AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" })
                            .ExecuteAsync();
                    return authResult.AccessToken;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Exception = ex;
                return null;
            }
        }

        /// <summary>
        /// Returns a Graph service client.
        /// </summary>
        /// <returns>GraphServiceClient</returns>
        public GraphServiceClient GetGraphClient()
        {
            try
            {
                return new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Token);
                }));
            }
            catch (Exception ex)
            {
                Exception = ex;
                return null;
            }
        }

        /// <summary>
        /// Load certificate file from directory.
        /// </summary>
        /// <param name="path">Path for certificate file.</param>
        /// <param name="password">Password for certificate file.</param>
        /// <returns>X509Certificate2</returns>
        private X509Certificate2 GetCertificateFromDirectory(string path, string password)
        {
            try
            {
                return new X509Certificate2(System.IO.Path.GetFullPath(path), password, X509KeyStorageFlags.MachineKeySet);
            }
            catch (Exception ex)
            {
                Exception = ex;
                return null;
            }
        }

        /// <summary>
        /// Load certificate file from directory.
        /// </summary>
        /// <param name="thumbprint">Certificate thumbprint.</param>
        /// <returns>X509Certificate2</returns>
        private X509Certificate2 GetCertificateFromStore(string thumbprint)
        {
            try
            {
                var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadOnly);
                var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                store.Close();
                return certificates[0];
            }
            catch (Exception ex)
            {
                Exception = ex;
                return null;
            }

        }
    }

    /// <summary>
    /// Type how to authorized in Azure.
    /// </summary>
    public enum AuthorizedBy
    {
        byClientSecret,
        byCertificatePath,
        byCertificateThumbprint
    }
}