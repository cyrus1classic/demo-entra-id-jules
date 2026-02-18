using EntraIDShowcase.API.Models.Licenses;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EntraIDShowcase.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = "Admin")]
    public class LicenseManagementController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public LicenseManagementController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/licensemanagement/subscriptions
        [HttpGet("subscriptions")]
        public async Task<IActionResult> GetSubscriptions()
        {
            var subscriptions = await _graphServiceClient.SubscribedSkus.GetAsync();

            if (subscriptions?.Value == null) return Ok(new List<SubscriptionDTO>());

            var subscriptionDtos = subscriptions.Value.Select(s => new SubscriptionDTO
            {
                SkuPartNumber = s.SkuPartNumber,
                PrepaidUnits = s.PrepaidUnits?.Enabled,
                ConsumedUnits = s.ConsumedUnits
            }).ToList();

            return Ok(subscriptionDtos);
        }

        // POST: api/licensemanagement/users/{userId}/assign-license
        [HttpPost("users/{userId}/assign-license")]
        public async Task<IActionResult> AssignLicense(string userId, [FromBody] AssignLicenseRequest request)
        {
             var assignLicenseBody = new Microsoft.Graph.Users.Item.AssignLicense.AssignLicensePostRequestBody
             {
                 AddLicenses = new List<AssignedLicense>
                 {
                     new AssignedLicense
                     {
                         SkuId = request.SkuId
                     }
                 },
                 RemoveLicenses = new List<Guid?>()
             };

             try
             {
                 await _graphServiceClient.Users[userId].AssignLicense.PostAsync(assignLicenseBody);
                 return Ok(new { message = "License assigned successfully" });
             }
             catch (ServiceException ex)
             {
                 return BadRequest(ex.Message);
             }
        }
    }
}
