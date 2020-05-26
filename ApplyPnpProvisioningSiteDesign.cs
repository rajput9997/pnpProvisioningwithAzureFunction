using System;
using System.Configuration;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace PnpProvisioningSiteDesign
{
    public static class ApplyPnpProvisioningSiteDesign
    {
        private static readonly string PNP_TEMPLATE_FILE = "ProjectTemplateV2Design.xml";

        [FunctionName("ApplyPnpProvisioningSiteDesign")]
        public static void Run([QueueTrigger("projectdesignqueue")]SiteInformation myQueueItem, TraceWriter log, ExecutionContext functionContext)
        {
            try
            {
                var clientID = ConfigurationManager.AppSettings["SPO_AppId"];
                var clientSecret = ConfigurationManager.AppSettings["SPO_AppSecret"];

                log.Info($"Fetched client ID '{clientID}' from AppSettings.'");

                ClientCredentials provisioningCreds = new ClientCredentials { ClientID = clientID, ClientSecret = clientSecret };
                applyTemplate(myQueueItem, functionContext, provisioningCreds, log);
            }
            catch (Exception e)
            {
                log.Error($"Error when running ApplyPnPTemplate function. Exception: {e}");
                throw;
            }
        }

        private static void applyTemplate(SiteInformation siteInformation, ExecutionContext functionContext, ClientCredentials credentials, TraceWriter log)
        {
            try
            {
                using (var ctx = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteInformation.SiteUrl, credentials.ClientID, credentials.ClientSecret))
                {
                    Web web = ctx.Web;
                    ctx.Load(web, w => w.Title);
                    ctx.ExecuteQueryRetry();

                    

                    var rootSiteUrl = ConfigurationManager.AppSettings["RootSiteUrl"];

                    log.Info($"Successfully connected to site: {web.Title}");

                    string currentDirectory = functionContext.FunctionDirectory;
                    DirectoryInfo dInfo = new DirectoryInfo(currentDirectory);
                    var schemaDir = dInfo.Parent.FullName + "\\Templates";
                    XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");

                    log.Info($"About to get template with with filename '{PNP_TEMPLATE_FILE}'");

                    ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);

                    var getHomeClientPage = template.ClientSidePages.Find(i => i.Title == "Home");
                    if (getHomeClientPage != null)
                    {
                        if (!string.IsNullOrWhiteSpace(siteInformation.ProposalDeadLineDate))
                        {
                            var proposalDate = Convert.ToDateTime(siteInformation.ProposalDeadLineDate);
                            siteInformation.ProposalDeadLineDate = string.Format("{0} of {1}, {2} {3}", proposalDate.Day.Ordinal(), proposalDate.ToString("MMMM"), proposalDate.ToString("yyyy"), proposalDate.ToString("hh:mm tt"));
                        }

                        var allSectionControls = getHomeClientPage.Sections?[0].Controls;

                        allSectionControls?[0].ControlProperties.Remove("Text");
                        allSectionControls?[0].ControlProperties.Add("Text", "<h4>Proposal Description -</h4><h4><span><span><span>" + siteInformation.Description + "</span></span></span></h4>");

                        allSectionControls?[1].ControlProperties.Remove("Text");
                        allSectionControls?[1].ControlProperties.Add("Text", "<h4>Proposal deadline -&nbsp; " + siteInformation.ProposalDeadLineDate + "&nbsp;</h4>");

                        if (!string.IsNullOrWhiteSpace(siteInformation.ProposalManager))
                        {
                            var managercoll = siteInformation.ProposalManager.Split(';');
                            var allPeoplesControls = getHomeClientPage.Sections?[2].Controls;

                            var jsondata = allPeoplesControls?[0].JsonControlData?.
                                Replace("jmerrell@umwelt.com.au", managercoll[1]).
                                Replace("John Merrell", managercoll[0]);
                            allPeoplesControls[0].JsonControlData = jsondata;
                        }
                    }

                    log.Info($"Successfully found template with ID '{template.Id}'");

                    ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                    {
                        ProgressDelegate = (message, progress, total) =>
                        {
                            log.Info(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                        }
                    };

                    // Associate file connector for assets..
                    FileSystemConnector connector = new FileSystemConnector(Path.Combine(currentDirectory, "Files"), "");
                    template.Connector = connector;

                    web.ApplyProvisioningTemplate(template, ptai);

                    // Add top navigation bar.
                    web.AddNavigationNode("SharePoint Main Menu", new Uri(rootSiteUrl + "/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    web.AddNavigationNode("Document Centre", new Uri(rootSiteUrl + "/Document%20Centre/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    web.AddNavigationNode("Project Centre", new Uri(rootSiteUrl + "/Project%20Centre/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    web.AddNavigationNode("WHS Centre", new Uri(rootSiteUrl + "/WHS%20Centre/"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    web.AddNavigationNode("Training Centre", new Uri(rootSiteUrl + "/Training%20Centre"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    web.AddNavigationNode("Proposal Hub", new Uri(rootSiteUrl + "/Proposal%20Hub"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);

                    UpdateListTitle(ctx, web, "01. Project Management");
                }
            }
            catch (Exception e)
            {
                log.Error("Error when applying PnP template!", e);
                throw;
            }
        }

        public static void UpdateListTitle(ClientContext ctx,Web web, string listTitle)
        {
            List odocumentList = web.Lists.GetByTitle("Documents");
            odocumentList.Title = listTitle;
            odocumentList.Update();
            ctx.ExecuteQuery();
        }

        public static string Ordinal(this int number)
        {
            var work = number.ToString();
            if ((number % 100) == 11 || (number % 100) == 12 || (number % 100) == 13)
                return work + "th";
            switch (number % 10)
            {
                case 1: work += "st"; break;
                case 2: work += "nd"; break;
                case 3: work += "rd"; break;
                default: work += "th"; break;
            }
            return work;
        }
    }
}
