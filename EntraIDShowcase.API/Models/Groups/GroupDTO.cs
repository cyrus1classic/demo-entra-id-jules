using System.Collections.Generic;

namespace EntraIDShowcase.API.Models.Groups
{
    public class GroupDTO
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public List<string> GroupTypes { get; set; }
        public bool? MailEnabled { get; set; }
        public bool? SecurityEnabled { get; set; }
    }
}
