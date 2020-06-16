using System;
using System.Configuration;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
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

                    // string groupID = GetSiteGroupID(ctx);
                    // UpdateSubscriptionItemProperties(credentials, siteInformation, web.Title ,log);
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
                        UpdateControlsDataDynamic(siteInformation, getHomeClientPage);
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

                    if (siteInformation.IsTopNavigation)
                    {
                        // Add top navigation bar.
                        web.AddNavigationNode("SharePoint Main Menu", new Uri(rootSiteUrl + "/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                        web.AddNavigationNode("Document Centre", new Uri(rootSiteUrl + "/Document%20Centre/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                        web.AddNavigationNode("Project Centre", new Uri(rootSiteUrl + "/Project%20Centre/SitePages/Home.aspx"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                        web.AddNavigationNode("WHS Centre", new Uri(rootSiteUrl + "/WHS%20Centre/"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                        web.AddNavigationNode("Training Centre", new Uri(rootSiteUrl + "/Training%20Centre"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                        web.AddNavigationNode("Proposal Hub", new Uri(rootSiteUrl + "/Proposal%20Hub"), "", OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
                    }

                    UpdateListTitle(ctx, web, "01. Project Management");
                }
            }
            catch (Exception e)
            {
                log.Error("Error when applying PnP template!", e);
                throw;
            }
        }

        public static void UpdateControlsDataDynamic(SiteInformation siteInformation, ClientSidePage getHomeClientPage)
        {
            string proposalDateTime = string.Empty;
            string proposalStartDate = string.Empty;
            if (!string.IsNullOrWhiteSpace(siteInformation.ProposalDeadLineDate))
            {
                var proposalDate = Convert.ToDateTime(siteInformation.ProposalDeadLineDate);
                proposalDateTime = string.Format("{0} of {1}, {2} {3}", proposalDate.Day.Ordinal(), proposalDate.ToString("MMMM"), proposalDate.ToString("yyyy"), proposalDate.ToString("hh:mm tt"));
            }
            if (!string.IsNullOrWhiteSpace(siteInformation.ProposalStartDate))
            {
                var proposalDate = Convert.ToDateTime(siteInformation.ProposalStartDate);
                proposalStartDate = string.Format("{0} of {1}, {2} {3}", proposalDate.Day.Ordinal(), proposalDate.ToString("MMMM"), proposalDate.ToString("yyyy"), proposalDate.ToString("hh:mm tt"));
            }

            siteInformation.ProposalStartDate = proposalStartDate;
            siteInformation.ProposalDeadLineDate = proposalDateTime;

            var allSectionControls = getHomeClientPage.Sections?[0].Controls;

            allSectionControls?[0].ControlProperties.Remove("Text");
            allSectionControls?[0].ControlProperties.Add("Text", "<h4>Project Description -</h4><h4><span><span><span>" + siteInformation.Description + "</span></span></span></h4>");

            allSectionControls?[1].ControlProperties.Remove("Text");
            allSectionControls?[1].ControlProperties.Add("Text", "<h4>Project Start Date -&nbsp; " + siteInformation.ProposalStartDate + "&nbsp;</h4>" +
                "                                        <br/> <h4>Delivery Date -&nbsp; " + siteInformation.ProposalDeadLineDate + "&nbsp;</h4>");

            var proposalmanagercontrol = getHomeClientPage.Sections?[2].Controls?[1];
            if (proposalmanagercontrol != null)
            {
                string managerEmailId = "";
                string managerDisplayName = "";
                if (!string.IsNullOrWhiteSpace(siteInformation.ProposalManager))
                {
                    var managercoll = siteInformation.ProposalManager.Split(';');
                    managerEmailId = managercoll[1];
                    managerDisplayName = managercoll[0];
                }
                var jsondata = proposalmanagercontrol.JsonControlData?.
                                    Replace("kdavies@umwelt.com.au", managerEmailId).
                                    Replace("kirsty davies", managerDisplayName);
                proposalmanagercontrol.JsonControlData = jsondata;
            }

            var proposaldirectorcontrol = getHomeClientPage.Sections?[2].Controls?[0];
            if (proposaldirectorcontrol != null)
            {
                string directorEmailId = "";
                string directorDisplayName = "";
                if (!string.IsNullOrWhiteSpace(siteInformation.ProposalDirector))
                {
                    var proposalcoll = siteInformation.ProposalDirector.Split(';');
                    directorEmailId = proposalcoll[1];
                    directorDisplayName = proposalcoll[0];
                }
                var jsondata = proposaldirectorcontrol.JsonControlData?.
                                    Replace("jmerrell@umwelt.com.au", directorEmailId).
                                    Replace("John Merrell", directorDisplayName);
                proposaldirectorcontrol.JsonControlData = jsondata;
            }

            var projectLocationControl = getHomeClientPage.Sections?[2].Controls?[2];
            if(projectLocationControl != null && !string.IsNullOrWhiteSpace(siteInformation.ProjectLocationName))
            {
                string jsondata = projectLocationControl.JsonControlData?.
                                    Replace("\"title\":\"\"", "\"title\": \"" + siteInformation.ProjectLocationName + "\"").
                                    Replace("\"defaultTitle\":\"\"", "\"defaultTitle\": \""+ siteInformation.ProjectLocationName +"\"").
                                    Replace("\"defaultAddress\":\"\"", "\"defaultAddress\": \"" + siteInformation.ProjectLocationAddress + "\"").
                                    Replace("\"address\":\"\"", "\"address\": \"" + siteInformation.ProjectLocationAddress + "\"");
                projectLocationControl.JsonControlData = jsondata;
            }
        }

        private static string GetSiteGroupID(ClientContext ctx)
        {
            try
            {
                var spSite = ctx.Site;
                ctx.Load(spSite, s => s.GroupId);
                ctx.ExecuteQuery();

                if (spSite.GroupId == new Guid())
                    return null;
                else
                    return spSite.GroupId.ToString();
            }
            catch(Exception ex) { return string.Empty; }
        }

        public static void UpdateSubscriptionItemProperties(ClientCredentials credentials, SiteInformation siteInformation, string webTitle, TraceWriter log)
        {
            try
            {
                var listTitle = ConfigurationManager.AppSettings["ProjectListTitle"];
                var projectSiteUrl = ConfigurationManager.AppSettings["ProjectSiteUrl"];
                using (var ctxProjectSite = new AuthenticationManager().GetAppOnlyAuthenticatedContext(projectSiteUrl, credentials.ClientID, credentials.ClientSecret))
                {
                    Web projectWeb = ctxProjectSite.Web;
                    ctxProjectSite.Load(projectWeb, w => w.Title);
                    ctxProjectSite.ExecuteQueryRetry();

                    List oList = projectWeb.Lists.GetByTitle(listTitle);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                        "<Value Type='Text'>" + webTitle + "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
                    ListItemCollection collListItem = oList.GetItems(camlQuery);

                    ctxProjectSite.Load(collListItem);
                    ctxProjectSite.ExecuteQueryRetry();

                    foreach (ListItem oListItem in collListItem)
                    {
                        siteInformation.Description = oListItem["Description"]?.ToString();
                        siteInformation.ProposalDirector = oListItem["PD_x002f_PM"]?.ToString();
                        siteInformation.ProposalDeadLineDate = oListItem["ProposalDeadline"]?.ToString();
                        siteInformation.ProposalStartDate = oListItem["ProjectStartDate"]?.ToString();
                        siteInformation.ProposalManager = oListItem["ProposalManager"]?.ToString();
                        siteInformation.ProjectLocationName = oListItem["ProjectLocationName"]?.ToString();
                        siteInformation.ProjectLocationAddress = oListItem["ProjectLocationAddress"]?.ToString();
                    }
                }
            }
            catch(Exception ex)
            {
                log.Error("Error when UpdateSubscriptionItemProperties", ex);
            }
        }

        public static void UpdateListTitle(ClientContext ctx, Web web, string listTitle)
        {
            try
            {
                List odocumentList = web.Lists.GetByTitle("Documents");
                odocumentList.Title = listTitle;
                odocumentList.Update();
                ctx.ExecuteQuery();
            }
            catch (Exception ex) { }
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
        /* public static async Task GetOffice365Group(Web web, Stream stream)
         {
             try
             {
                 string clientId = "b7ec8c67-495f-4b8c-8c91-442147c480af";
                 string clientSecret = "45-A~v-U_-zFi5LJ.I.O91__em0-IrFAgb"; //ConfigurationManager.AppSettings["ClientSecret"];
                 string authority = "https://login.microsoftonline.com/927c0756-779d-44e0-baa5-8f0ea58bd36e/oauth2/v2.0/authorize"; //ConfigurationManager.AppSettings["AuthorityUrl"];

                 AuthenticationContext authContext = new AuthenticationContext(authority);
                 ClientCredential creds = new ClientCredential(clientId, clientSecret);
                 AuthenticationResult authResult = await authContext.AcquireTokenAsync("https://graph.microsoft.com/", creds);
                 string token = authResult.AccessToken;

                 Stream groupLogoStream = new FileStream("C:\\Jignesh\\Umwelt_ProjectSite_Logo.png",
                                            FileMode.Open, FileAccess.Read);
                 var groupColl = UnifiedGroupsUtility.ListUnifiedGroups(token, web.Title);
                 if (groupColl.Count > 0)
                 {
                     var ocurrentGroup = groupColl[0];
                     UnifiedGroupsUtility.UpdateUnifiedGroup(ocurrentGroup.GroupId, token, description: "my test",
                                           groupLogo: groupLogoStream );
                 }
             }
             catch (Exception ex)
             {
                 string b = ex.Message;
             }
         } */
    }
}
