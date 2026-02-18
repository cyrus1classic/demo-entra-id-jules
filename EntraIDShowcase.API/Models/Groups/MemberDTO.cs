namespace EntraIDShowcase.API.Models.Groups
{
    public class MemberDTO
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string UserPrincipalName { get; set; }
        public string Type { get; set; } // "User" or "Group"
    }
}
