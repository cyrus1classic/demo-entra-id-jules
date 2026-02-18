using EntraIDShowcase.API.Models.Auth;
using EntraIDShowcase.API.Models.Users;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace EntraIDShowcase.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthenticationController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IConfiguration _configuration;

        public AuthenticationController(GraphServiceClient graphServiceClient, IConfiguration configuration)
        {
            _graphServiceClient = graphServiceClient;
            _configuration = configuration;
        }

        // POST: api/authentication/login
        [HttpPost("login")]
        [AllowAnonymous]
        public async Task<IActionResult> Login([FromBody] LoginRequest request)
        {
            // ROPC Flow - For demonstration purposes only. Not recommended for production.
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var instance = _configuration["AzureAd:Instance"] ?? "https://login.microsoftonline.com/";

            // For the API to accept the token, the audience must be the API itself.
            // Typically, the scope is "api://<ClientId>/.default" or similar.
            // If the App Registration exposes an API, use that scope.
            // If not, we can try using the ClientId as the scope for the token to be issued for the app itself (if allowed).
            var scopes = new[] { $"{clientId}/.default" };

            try
            {
                // Note: ROPC flow is not supported for accounts with MFA enabled or Federated accounts.
                // This will fail if the user has MFA.
                // In a real app, use Authorization Code Flow (frontend redirects user).

                // Using PublicClientApplication for ROPC demonstration.
                // This simulates a public client (like a mobile app) authenticating against Entra ID to get a token for this API.
                var publicApp = PublicClientApplicationBuilder.Create(clientId)
                    .WithAuthority($"{instance}{tenantId}")
                    .Build();

                var result = await publicApp.AcquireTokenByUsernamePassword(scopes, request.Username, request.Password)
                    .ExecuteAsync();

                return Ok(new
                {
                    accessToken = result.AccessToken,
                    expiresOn = result.ExpiresOn,
                    user = new { username = result.Account.Username }
                });
            }
            catch (MsalException ex)
            {
                return BadRequest(new { message = "Authentication failed", details = ex.Message });
            }
        }

        // POST: api/authentication/logout
        [HttpPost("logout")]
        [Authorize]
        public async Task<IActionResult> Logout()
        {
            // In a stateless API, "logout" is mostly client-side (clearing tokens).
            // We can return the sign-out URL.
            var tenantId = _configuration["AzureAd:TenantId"];
            var instance = _configuration["AzureAd:Instance"] ?? "https://login.microsoftonline.com/";

            var logoutUrl = $"{instance}{tenantId}/oauth2/v2.0/logout";

            return Ok(new { message = "Clear your local storage/cookies.", logoutUrl = logoutUrl });
        }

        // GET: api/authentication/me
        [HttpGet("me")]
        [Authorize]
        public async Task<IActionResult> GetCurrentUser()
        {
            try
            {
                // Uses the On-Behalf-Of flow or the incoming token to call Graph
                var me = await _graphServiceClient.Me.GetAsync();

                var userDto = new UserDTO
                {
                    Id = me.Id,
                    DisplayName = me.DisplayName,
                    Mail = me.Mail,
                    UserPrincipalName = me.UserPrincipalName,
                    Department = me.Department,
                    JobTitle = me.JobTitle
                };

                return Ok(userDto);
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // POST: api/authentication/mfa/enable
        [HttpPost("mfa/enable")]
        [Authorize]
        public async Task<IActionResult> EnableMFA()
        {
            // Enabling MFA is typically done via Conditional Access Policies or Per-user MFA settings (legacy).
            // Here we can showcase listing Authentication Methods to see if MFA is registered.

            try
            {
                 // Check registered authentication methods
                 // Note: Requires UserAuthenticationMethod.Read.All permission
                 var methods = await _graphServiceClient.Me.Authentication.Methods.GetAsync();

                 return Ok(new
                 {
                     message = "MFA cannot be directly 'enabled' via API for a single user without policy changes. Here are registered methods.",
                     methods = methods.Value.Select(m => m.OdataType).ToList()
                 });
            }
             catch (ServiceException ex)
            {
                return BadRequest(new { message = "Could not retrieve auth methods.", details = ex.Message });
            }
        }
    }
}
