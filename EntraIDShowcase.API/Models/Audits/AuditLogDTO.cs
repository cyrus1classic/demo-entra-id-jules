using System;

namespace EntraIDShowcase.API.Models.Audits
{
    public class AuditLogDTO
    {
        public DateTimeOffset? ActivityDateTime { get; set; }
        public string ActivityDisplayName { get; set; }
        public string InitiatedBy { get; set; }
        public string Result { get; set; }
    }
}
