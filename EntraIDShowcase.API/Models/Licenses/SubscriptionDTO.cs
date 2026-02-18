namespace EntraIDShowcase.API.Models.Licenses
{
    public class SubscriptionDTO
    {
        public string SkuPartNumber { get; set; }
        public int? PrepaidUnits { get; set; }
        public int? ConsumedUnits { get; set; }
    }
}
