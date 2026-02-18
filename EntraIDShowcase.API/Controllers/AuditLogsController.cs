using EntraIDShowcase.API.Models.Audits;
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
    [Authorize(Roles = "Admin,SecurityReader")]
    public class AuditLogsController : ControllerBase
    {
        private readonly GraphServiceClient _graphServiceClient;

        public AuditLogsController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: api/auditlogs/directory-audits
        [HttpGet("directory-audits")]
        public async Task<IActionResult> GetDirectoryAudits([FromQuery] DateTime? startDate, [FromQuery] DateTime? endDate)
        {
            try
            {
                var audits = await _graphServiceClient.AuditLogs.DirectoryAudits.GetAsync(requestConfiguration =>
                {
                     var filters = new List<string>();
                     if (startDate.HasValue) filters.Add($"activityDateTime ge {startDate.Value:yyyy-MM-ddTHH:mm:ssZ}");
                     if (endDate.HasValue) filters.Add($"activityDateTime le {endDate.Value:yyyy-MM-ddTHH:mm:ssZ}");

                     if (filters.Any())
                     {
                         requestConfiguration.QueryParameters.Filter = string.Join(" and ", filters);
                     }
                     requestConfiguration.QueryParameters.Top = 50;
                });

                if (audits?.Value == null) return Ok(new List<AuditLogDTO>());

                var auditDtos = audits.Value.Select(a => new AuditLogDTO
                {
                    ActivityDateTime = a.ActivityDateTime,
                    ActivityDisplayName = a.ActivityDisplayName,
                    InitiatedBy = a.InitiatedBy?.User?.DisplayName ?? a.InitiatedBy?.App?.DisplayName ?? "System",
                    Result = a.ResultReason ?? a.Result.ToString()
                }).ToList();

                return Ok(auditDtos);
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // GET: api/auditlogs/sign-ins
        [HttpGet("sign-ins")]
        public async Task<IActionResult> GetSignInLogs([FromQuery] string? userId, [FromQuery] DateTime? startDate)
        {
            try
            {
                var signIns = await _graphServiceClient.AuditLogs.SignIns.GetAsync(requestConfiguration =>
                {
                    var filters = new List<string>();
                    if (!string.IsNullOrEmpty(userId)) filters.Add($"userId eq '{userId}'");
                    if (startDate.HasValue) filters.Add($"createdDateTime ge {startDate.Value:yyyy-MM-ddTHH:mm:ssZ}");

                    if (filters.Any())
                    {
                        requestConfiguration.QueryParameters.Filter = string.Join(" and ", filters);
                    }
                    requestConfiguration.QueryParameters.Top = 50;
                });

                if (signIns?.Value == null) return Ok(new List<SignInDTO>());

                var signInDtos = signIns.Value.Select(s => new SignInDTO
                {
                    CreatedDateTime = s.CreatedDateTime,
                    UserDisplayName = s.UserDisplayName,
                    IpAddress = s.IpAddress,
                    Location = $"{s.Location?.City}, {s.Location?.CountryOrRegion}",
                    Status = s.Status?.ErrorCode == 0 ? "Success" : $"Failure: {s.Status?.FailureReason}"
                }).ToList();

                return Ok(signInDtos);
            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
