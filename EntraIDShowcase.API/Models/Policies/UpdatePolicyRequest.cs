using Microsoft.Graph.Models;

namespace EntraIDShowcase.API.Models.Policies
{
    public class UpdatePolicyRequest
    {
        public string DisplayName { get; set; }
        public string State { get; set; }
        public ConditionalAccessConditionSet Conditions { get; set; }
        public ConditionalAccessGrantControls GrantControls { get; set; }
    }
}
