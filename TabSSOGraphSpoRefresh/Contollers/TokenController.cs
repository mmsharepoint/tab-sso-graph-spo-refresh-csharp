using Azure.Identity;
using Microsoft.Graph.Models;
using Microsoft.SharePoint.Authentication;
using Microsoft.SharePoint.Client;

namespace TabSSOGraphSpoRefresh.Contollers
{
    public class TokenController
    {
        private readonly IConfiguration _configuration;
        public TokenController() 
        {
            //_configuration = configuration;
        }

        public async Task<string> getSiteUser(string clientId, string clientSecret, string tenantId, string userAssertion, string siteUrl, string userLogin)
        {
            string token = await getSPOToken(clientId, clientSecret, tenantId, userAssertion);
            ClientContext context = new ClientContext(siteUrl);

            context.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
            };

            var web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            return web.Title;
        }
        private async Task<string> getOBOToken(string clientId, string clientSecret, string tenantId, string userAssertion)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default", "offline_access" }; 

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            //var tenantId = _configuration.GetValue<string>("TeamsFx.Authentication.TenantId");

            // Values from app registration
            //var clientId = _configuration.GetValue<string>("TeamsFx.Authentication.ClientId");
            //var clientSecret = _configuration.GetValue<string>("TeamsFx.Authentication.ClientSecret");

            // using Azure.Identity;
            var options = new OnBehalfOfCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // This is the incoming token to exchange using on-behalf-of flow
            var oboToken = userAssertion;

            var onBehalfOfCredential = new OnBehalfOfCredential(
                tenantId, clientId, clientSecret, oboToken, options);

            var onBehalfO = await onBehalfOfCredential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), new CancellationToken());
            return onBehalfO.Token;
        }

        private async Task<string> getSPOToken(string clientId, string clientSecret, string tenantId, string userAssertion)
        {
            var scopes = new[] { "https://mmoeller.sharepoint.com/AllSites.Write" };

            // using Azure.Identity;
            var options = new OnBehalfOfCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // This is the incoming token to exchange using on-behalf-of flow
            var oboToken = userAssertion;

            var onBehalfOfCredential = new OnBehalfOfCredential(
                tenantId, clientId, clientSecret, oboToken, options);

            var onBehalfO = await onBehalfOfCredential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), new CancellationToken());
            return onBehalfO.Token;
        }
    }
}
