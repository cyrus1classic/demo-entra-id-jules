using EntraIDShowcase.API.Models.Groups;
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
    [Authorize]
    public class GroupManagementController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public GroupManagementController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/groupmanagement/groups
        [HttpGet("groups")]
        [Authorize(Roles = "Admin,GroupReader")]
        public async Task<IActionResult> GetAllGroups([FromQuery] string? filter)
        {
            var groups = await _graphServiceClient.Groups.GetAsync(requestConfiguration =>
            {
                if (!string.IsNullOrEmpty(filter))
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                }
                requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "description", "groupTypes", "mailEnabled", "securityEnabled" };
            });

            if (groups?.Value == null) return Ok(new List<GroupDTO>());

            var groupDtos = groups.Value.Select(g => new GroupDTO
            {
                Id = g.Id,
                DisplayName = g.DisplayName,
                Description = g.Description,
                GroupTypes = g.GroupTypes,
                MailEnabled = g.MailEnabled,
                SecurityEnabled = g.SecurityEnabled
            }).ToList();

            return Ok(groupDtos);
        }

        // GET: api/groupmanagement/groups/{groupId}/members
        [HttpGet("groups/{groupId}/members")]
        [Authorize(Roles = "Admin,GroupReader")]
        public async Task<IActionResult> GetGroupMembers(string groupId)
        {
            try
            {
                var members = await _graphServiceClient.Groups[groupId].Members.GetAsync();

                if (members?.Value == null) return Ok(new List<MemberDTO>());

                var memberDtos = members.Value.Select(m => {
                    var memberDto = new MemberDTO
                    {
                        Id = m.Id,
                        Type = m.OdataType.Replace("#microsoft.graph.", "")
                    };

                    if (m is User user)
                    {
                        memberDto.DisplayName = user.DisplayName;
                        memberDto.UserPrincipalName = user.UserPrincipalName;
                    }
                    else if (m is Group group)
                    {
                        memberDto.DisplayName = group.DisplayName;
                    }

                    return memberDto;
                }).ToList();

                return Ok(memberDtos);
            }
            catch (ServiceException ex) when (ex.ResponseStatusCode == 404)
            {
                return NotFound();
            }
        }

        // POST: api/groupmanagement/groups
        [HttpPost("groups")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> CreateGroup([FromBody] CreateGroupRequest request)
        {
            var group = new Group
            {
                DisplayName = request.DisplayName,
                Description = request.Description,
                MailNickname = request.MailNickname,
                GroupTypes = request.GroupTypes,
                MailEnabled = request.MailEnabled,
                SecurityEnabled = request.SecurityEnabled
            };

            try
            {
                var createdGroup = await _graphServiceClient.Groups.PostAsync(group);
                // Log audit event
                return CreatedAtAction(nameof(GetGroupMembers), new { groupId = createdGroup.Id }, new { id = createdGroup.Id });
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // POST: api/groupmanagement/groups/{groupId}/members
        [HttpPost("groups/{groupId}/members")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> AddMemberToGroup(string groupId, [FromBody] string userId)
        {
            var requestBody = new ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
            };

            try
            {
                await _graphServiceClient.Groups[groupId].Members.Ref.PostAsync(requestBody);
                // Log audit event
                return NoContent();
            }
            catch (ServiceException ex)
            {
                if (ex.ResponseStatusCode == 404) return NotFound();
                return BadRequest(ex.Message);
            }
        }

        // DELETE: api/groupmanagement/groups/{groupId}/members/{userId}
        [HttpDelete("groups/{groupId}/members/{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> RemoveMemberFromGroup(string groupId, string userId)
        {
            try
            {
                await _graphServiceClient.Groups[groupId].Members[userId].Ref.DeleteAsync();
                // Log audit event
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
