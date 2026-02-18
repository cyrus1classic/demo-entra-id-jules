using System.Collections.Generic;

namespace EntraIDShowcase.API.Models.Groups
{
    public class CreateGroupRequest
    {
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string MailNickname { get; set; }
        public List<string> GroupTypes { get; set; }
        public bool MailEnabled { get; set; }
        public bool SecurityEnabled { get; set; }
    }
}
