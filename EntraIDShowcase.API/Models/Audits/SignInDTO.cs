using System;

namespace EntraIDShowcase.API.Models.Audits
{
    public class SignInDTO
    {
        public DateTimeOffset? CreatedDateTime { get; set; }
        public string UserDisplayName { get; set; }
        public string IpAddress { get; set; }
        public string Location { get; set; }
        public string Status { get; set; }
    }
}
