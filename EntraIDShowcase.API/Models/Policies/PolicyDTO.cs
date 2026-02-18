using System.Collections.Generic;

namespace EntraIDShowcase.API.Models.Policies
{
    public class PolicyDTO
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string State { get; set; }
        public object Conditions { get; set; } // Simplified for showcase
        public object GrantControls { get; set; } // Simplified for showcase
    }
}
