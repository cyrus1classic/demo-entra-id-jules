using EntraIDShowcase.API.Models.Policies;
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
    [Authorize(Roles = "Admin,SecurityAdmin")]
    public class ConditionalAccessController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public ConditionalAccessController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/conditionalaccess/policies
        [HttpGet("policies")]
        public async Task<IActionResult> GetPolicies()
        {
            var policies = await _graphServiceClient.Identity.ConditionalAccess.Policies.GetAsync();

            if (policies?.Value == null) return Ok(new List<PolicyDTO>());

            var policyDtos = policies.Value.Select(p => new PolicyDTO
            {
                Id = p.Id,
                DisplayName = p.DisplayName,
                State = p.State.ToString(),
                Conditions = p.Conditions,
                GrantControls = p.GrantControls
            }).ToList();

            return Ok(policyDtos);
        }

        // GET: api/conditionalaccess/policies/{policyId}
        [HttpGet("policies/{policyId}")]
        public async Task<IActionResult> GetPolicyById(string policyId)
        {
            try
            {
                var policy = await _graphServiceClient.Identity.ConditionalAccess.Policies[policyId].GetAsync();

                if (policy == null) return NotFound();

                var policyDto = new PolicyDTO
                {
                    Id = policy.Id,
                    DisplayName = policy.DisplayName,
                    State = policy.State.ToString(),
                    Conditions = policy.Conditions,
                    GrantControls = policy.GrantControls
                };

                return Ok(policyDto);
            }
             catch (ServiceException ex) when (ex.ResponseStatusCode == 404)
            {
                return NotFound();
            }
        }

        // POST: api/conditionalaccess/policies
        [HttpPost("policies")]
        public async Task<IActionResult> CreatePolicy([FromBody] CreatePolicyRequest request)
        {
            if (!Enum.TryParse<ConditionalAccessPolicyState>(request.State, true, out var state))
            {
                return BadRequest("Invalid policy state.");
            }

            var policy = new ConditionalAccessPolicy
            {
                DisplayName = request.DisplayName,
                State = state,
                Conditions = request.Conditions,
                GrantControls = request.GrantControls
            };

            try
            {
                var createdPolicy = await _graphServiceClient.Identity.ConditionalAccess.Policies.PostAsync(policy);
                return CreatedAtAction(nameof(GetPolicyById), new { policyId = createdPolicy.Id }, new { id = createdPolicy.Id });
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // PATCH: api/conditionalaccess/policies/{policyId}
        [HttpPatch("policies/{policyId}")]
        public async Task<IActionResult> UpdatePolicy(string policyId, [FromBody] UpdatePolicyRequest request)
        {
             if (!Enum.TryParse<ConditionalAccessPolicyState>(request.State, true, out var state))
             {
                 return BadRequest("Invalid policy state.");
             }

             var policy = new ConditionalAccessPolicy
             {
                 DisplayName = request.DisplayName,
                 State = state,
                 Conditions = request.Conditions,
                 GrantControls = request.GrantControls
             };

             try
             {
                 await _graphServiceClient.Identity.ConditionalAccess.Policies[policyId].PatchAsync(policy);
                 return NoContent();
             }
             catch (ServiceException ex)
             {
                 if (ex.ResponseStatusCode == 404) return NotFound();
                 return BadRequest(ex.Message);
             }
        }
    }
}
