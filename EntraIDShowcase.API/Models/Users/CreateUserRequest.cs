namespace EntraIDShowcase.API.Models.Users
{
    public class CreateUserRequest
    {
        public string DisplayName { get; set; }
        public string MailNickname { get; set; }
        public string UserPrincipalName { get; set; }
        public string Password { get; set; }
        public bool ForceChangePasswordNextSignIn { get; set; }
        public bool AccountEnabled { get; set; }
    }
}
