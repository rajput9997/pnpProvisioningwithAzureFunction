namespace PnpProvisioningSiteDesign
{
    public class ClientCredentials
    {
        public string ClientID { get; set; }
        public string ClientSecret { get; set; }
    }

    public class SiteInformation
    {
        public string SiteUrl { get; set; }
        public string Description { get; set; }
        public string ProposalStartDate { get; set; }
        public string ProposalDeadLineDate { get; set; }
        public string ProposalManager { get; set; }
        public string ProposalDirector { get; set; }
    }
}
