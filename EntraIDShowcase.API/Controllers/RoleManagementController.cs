using EntraIDShowcase.API.Models.Roles;
using EntraIDShowcase.API.Models.Users;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EntraIDShowcase.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = "Admin")]
    public class RoleManagementController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public RoleManagementController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/rolemanagement/directory-roles
        [HttpGet("directory-roles")]
        public async Task<IActionResult> GetDirectoryRoles()
        {
            var roles = await _graphServiceClient.DirectoryRoles.GetAsync();

            if (roles?.Value == null) return Ok(new List<RoleDTO>());

            var roleDtos = roles.Value.Select(r => new RoleDTO
            {
                Id = r.Id,
                DisplayName = r.DisplayName,
                Description = r.Description
            }).ToList();

            return Ok(roleDtos);
        }

        // GET: api/rolemanagement/directory-roles/{roleId}/members
        [HttpGet("directory-roles/{roleId}/members")]
        public async Task<IActionResult> GetRoleMembers(string roleId)
        {
             try
             {
                 var members = await _graphServiceClient.DirectoryRoles[roleId].Members.GetAsync();

                 if (members?.Value == null) return Ok(new List<UserDTO>());

                 var userDtos = members.Value.OfType<User>().Select(u => new UserDTO
                 {
                     Id = u.Id,
                     DisplayName = u.DisplayName,
                     Mail = u.Mail,
                     UserPrincipalName = u.UserPrincipalName,
                     Department = u.Department,
                     JobTitle = u.JobTitle
                 }).ToList();

                 return Ok(userDtos);
             }
             catch (ServiceException ex) when (ex.ResponseStatusCode == 404)
             {
                 return NotFound();
             }
        }

        // POST: api/rolemanagement/directory-roles/{roleId}/members
        [HttpPost("directory-roles/{roleId}/members")]
        public async Task<IActionResult> AssignRoleToUser(string roleId, [FromBody] string userId)
        {
             var requestBody = new ReferenceCreate
             {
                 OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
             };

             try
             {
                 await _graphServiceClient.DirectoryRoles[roleId].Members.Ref.PostAsync(requestBody);
                 // Log audit event
                 return NoContent();
             }
             catch (ServiceException ex)
             {
                 if (ex.ResponseStatusCode == 404) return NotFound();
                 return BadRequest(ex.Message);
             }
        }

        // GET: api/rolemanagement/app-roles
        [HttpGet("app-roles")]
        public async Task<IActionResult> GetApplicationRoles([FromQuery] string appId)
        {
             var servicePrincipals = await _graphServiceClient.ServicePrincipals.GetAsync(requestConfiguration =>
             {
                 requestConfiguration.QueryParameters.Filter = $"appId eq '{appId}'";
             });

             var sp = servicePrincipals?.Value?.FirstOrDefault();
             if (sp == null) return NotFound("Application not found");

             var appRoles = sp.AppRoles?.Select(r => new RoleDTO
             {
                 Id = r.Id.ToString(),
                 DisplayName = r.DisplayName,
                 Description = r.Description
             }).ToList();

             return Ok(appRoles ?? new List<RoleDTO>());
        }
    }
}
