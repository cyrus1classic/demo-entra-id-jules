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
    [Authorize]
    public class UserManagementController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public UserManagementController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/usermanagement/users
        [HttpGet("users")]
        [Authorize(Roles = "Admin,UserReader")]
        public async Task<IActionResult> GetAllUsers([FromQuery] string? filter, [FromQuery] int pageSize = 50)
        {
            var users = await _graphServiceClient.Users.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = pageSize;
                if (!string.IsNullOrEmpty(filter))
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                }
                requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "mail", "userPrincipalName", "department", "jobTitle" };
            });

            if (users?.Value == null) return Ok(new List<UserDTO>());

            var userDtos = users.Value.Select(u => new UserDTO
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

        // GET: api/usermanagement/users/{userId}
        [HttpGet("users/{userId}")]
        [Authorize(Roles = "Admin,UserReader")]
        public async Task<IActionResult> GetUserById(string userId)
        {
            try
            {
                var user = await _graphServiceClient.Users[userId].GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "mail", "userPrincipalName", "department", "jobTitle", "assignedLicenses" };
                    requestConfiguration.QueryParameters.Expand = new[] { "manager", "memberOf" };
                });

                if (user == null) return NotFound();

                var userDto = new UserDTO
                {
                    Id = user.Id,
                    DisplayName = user.DisplayName,
                    Mail = user.Mail,
                    UserPrincipalName = user.UserPrincipalName,
                    Department = user.Department,
                    JobTitle = user.JobTitle
                };

                return Ok(userDto);
            }
            catch (ServiceException ex) when (ex.ResponseStatusCode == 404)
            {
                return NotFound();
            }
        }

        // POST: api/usermanagement/users
        [HttpPost("users")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> CreateUser([FromBody] CreateUserRequest request)
        {
            var user = new User
            {
                AccountEnabled = request.AccountEnabled,
                DisplayName = request.DisplayName,
                MailNickname = request.MailNickname,
                UserPrincipalName = request.UserPrincipalName,
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = request.ForceChangePasswordNextSignIn,
                    Password = request.Password
                }
            };

            try
            {
                var createdUser = await _graphServiceClient.Users.PostAsync(user);
                // Log audit event (omitted for brevity)
                return CreatedAtAction(nameof(GetUserById), new { userId = createdUser.Id }, new { id = createdUser.Id });
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // PATCH: api/usermanagement/users/{userId}
        [HttpPatch("users/{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> UpdateUser(string userId, [FromBody] UpdateUserRequest request)
        {
            var user = new User
            {
                DisplayName = request.DisplayName,
                JobTitle = request.JobTitle,
                Department = request.Department
            };

            try
            {
                await _graphServiceClient.Users[userId].PatchAsync(user);
                // Log audit event
                return NoContent();
            }
            catch (ServiceException ex)
            {
                 if (ex.ResponseStatusCode == 404) return NotFound();
                 return BadRequest(ex.Message);
            }
        }

        // DELETE: api/usermanagement/users/{userId}
        [HttpDelete("users/{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteUser(string userId)
        {
            try
            {
                await _graphServiceClient.Users[userId].DeleteAsync();
                // Log audit event
                return NoContent();
            }
            catch (ServiceException ex)
            {
                if (ex.ResponseStatusCode == 404) return NotFound();
                return BadRequest(ex.Message);
            }
        }

        // POST: api/usermanagement/users/{userId}/reset-password
        [HttpPost("users/{userId}/reset-password")]
        [Authorize(Roles = "Admin,PasswordAdmin")]
        public async Task<IActionResult> ResetPassword(string userId, [FromBody] ResetPasswordRequest request)
        {
             var user = new User
             {
                 PasswordProfile = new PasswordProfile
                 {
                     ForceChangePasswordNextSignIn = true,
                     Password = request.NewPassword
                 }
             };

             try
             {
                 await _graphServiceClient.Users[userId].PatchAsync(user);
                 // Log audit event
                 return NoContent();
             }
             catch (ServiceException ex)
             {
                 if (ex.ResponseStatusCode == 404) return NotFound();
                 return BadRequest(ex.Message);
             }
        }

        // POST: api/usermanagement/users/{userId}/enable-disable
        [HttpPost("users/{userId}/enable-disable")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> ToggleUserStatus(string userId, [FromBody] bool enable)
        {
             var user = new User
             {
                 AccountEnabled = enable
             };

             try
             {
                 await _graphServiceClient.Users[userId].PatchAsync(user);
                 // Log audit event
                 return Ok(new { accountEnabled = enable });
             }
             catch (ServiceException ex)
             {
                 if (ex.ResponseStatusCode == 404) return NotFound();
                 return BadRequest(ex.Message);
             }
        }
    }
}
