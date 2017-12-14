using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.PowerBI.Api.V2;
using Microsoft.Rest;
using System;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using PowerBIEmbedded_Native.types;

namespace PowerBIEmbedded_Native
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The Tenant is the name of the Azure AD tenant in which this application is registered.
        // The AAD Instance is the instance of Azure, for example public Azure or Azure China.
        // The Redirect URI is the URI where Azure AD will return OAuth responses.
        // The Authority is the sign-in URL of the tenant.
        //
        private static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static string tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        Uri redirectUri = new Uri(ConfigurationManager.AppSettings["ida:RedirectUri"]);
        private static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);
        
        private static string graphResourceId = ConfigurationManager.AppSettings["ida:ResourceId"];
        private AuthenticationContext authContext = null;

        TokenCredentials tokenCredentials = null;
        string Token = null;
        string ApiUrl = "https://api.powerbi.com";



        public MainWindow()
        {
            InitializeComponent();
            TokenCache TC = new TokenCache();
            authContext = new AuthenticationContext(authority, TC);
        }


        private void getAppWorkspacesList()
        {
            using (var client = new PowerBIClient(new Uri(ApiUrl), tokenCredentials))
            {
                appWorkSpacesList.ItemsSource = client.Groups.GetGroups().Value.Select(g => new workSpaceList(g.Name, g.Id));
            }
        }

        private async void LoginAAD_Click(object sender, RoutedEventArgs e)
        {

            AuthenticationResult result = null;

            try
            {
                result = await authContext.AcquireTokenAsync(graphResourceId, clientId, redirectUri, new PlatformParameters(PromptBehavior.SelectAccount));
                Token = result.AccessToken;
                tokenCredentials = new TokenCredentials(Token, "Bearer");
                getAppWorkspacesList();

            }
            catch (AdalException ex)
            {
                // An unexpected error occurred, or user canceled the sign in.
                if (ex.ErrorCode != "access_denied")
                    MessageBox.Show(ex.Message);

                return;
            }
        }

        private void Embed_Click(object sender, RoutedEventArgs e)
        {
            Uri uri = new Uri(@"pack://siteoforigin:,,,/html/ReportLoader.html");
            PBIEmbeddedWB.ObjectForScripting = new AddJavascriptObjects(EmbeddedLogger, PBIEmbedded_Invoke);
            PBIEmbeddedWB.Navigate(uri);
        }

        public void HideScriptErrors(WebBrowser wb, bool Hide)
        {
            FieldInfo fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            object objComWebBrowser = fiComWebBrowser.GetValue(wb);
            if (objComWebBrowser == null) return;
            objComWebBrowser.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { Hide });
        }

        public void PBIEmbedded_Invoke()
        {
            PBIContentObject embed = null;
            object[] parameters = null;


            if (DashboardSelected.IsChecked.Value && appWorkSpacesDashboardList.SelectedItem != null)
            {
                embed = (PBIContentObject)appWorkSpacesDashboardList.SelectedItem;
                parameters = new object[] { embed.EmbeddedUrl, Token, embed.EmbeddedId, PBIObjectType.Dashboard, 0, string.Empty };
                PBIEmbeddedWB.InvokeScript("LoadEmbeddedObject", parameters);
            }
            else if (ReportSelected.IsChecked.Value && appWorkSpacesReportList.SelectedItem != null)
            {
                embed = (PBIContentObject)appWorkSpacesReportList.SelectedItem;
                parameters = new object[] { embed.EmbeddedUrl, Token, embed.EmbeddedId, PBIObjectType.Report, 0, string.Empty };
                PBIEmbeddedWB.InvokeScript("LoadEmbeddedObject", parameters);
            }
            else if (TileSelected.IsChecked.Value && appWorkSpacesDashboardList.SelectedItem != null & appWorkSpacesTileList.SelectedItem != null)
            {
                string dashboardId = ((PBIContentObject)appWorkSpacesDashboardList.SelectedItem).EmbeddedId;
                embed = (PBIContentObject)appWorkSpacesTileList.SelectedItem;
                parameters = new object[] { embed.EmbeddedUrl, Token, embed.EmbeddedId, PBIObjectType.Tile, 0, dashboardId };
                PBIEmbeddedWB.InvokeScript("LoadEmbeddedObject", parameters);
            }

        }

        private void PBIEmbeddedWB_Navigated(object sender, NavigationEventArgs e)
        {
            HideScriptErrors(PBIEmbeddedWB, true);

        }

        [ComVisible(true)]
        public class AddJavascriptObjects
        {
            TextBlock logWindow = null;
            Action fireJavaScript = null;
            public AddJavascriptObjects(TextBlock logWindow, Action action)
            {
                this.logWindow = logWindow;
                this.fireJavaScript = action;
            }
            public void LogToBrowserHost(string message)
            {
                logWindow.Text = logWindow.Text + message + "\r\n";
                // Force logger to bottom of scroll
                var parentContainer = (ScrollViewer)logWindow.Parent;
                parentContainer.UpdateLayout();
                parentContainer.ScrollToVerticalOffset(parentContainer.ScrollableHeight);

            }
            public void triggerDocumentComplete()
            {
                fireJavaScript();
            }
        }

        private void appWorkSpacesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (appWorkSpacesList.SelectedItem != null)
            {
                using (var client = new PowerBIClient(new Uri(ApiUrl), tokenCredentials))
                {
                    var WorkSpace = appWorkSpacesList.SelectedItem as workSpaceList;

                    appWorkSpacesDashboardList.ItemsSource = client.Dashboards.GetDashboardsInGroup(WorkSpace.Id).Value.Select(d => new PBIContentObject(d.DisplayName, d.EmbedUrl, d.Id));
                    appWorkSpacesDashboardList.SelectedIndex = 0;

                    appWorkSpacesReportList.ItemsSource = client.Reports.GetReportsInGroup(WorkSpace.Id).Value.Select(r => new PBIContentObject(r.Name, r.EmbedUrl, r.Id));
                    appWorkSpacesReportList.SelectedIndex = 0;
                }
            }

        }

        private void appWorkSpacesDashboardList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (appWorkSpacesDashboardList.SelectedItem != null)
            {
                using (var client = new PowerBIClient(new Uri(ApiUrl), tokenCredentials))
                {
                    var WorkSpace = appWorkSpacesList.SelectedItem as workSpaceList;
                    PBIContentObject dashboard = (PBIContentObject)appWorkSpacesDashboardList.SelectedItem;
                    appWorkSpacesTileList.ItemsSource = client.Dashboards.GetTilesInGroup(WorkSpace.Id, dashboard.EmbeddedId).Value.Select(t => new PBIContentObject(t.Title, t.EmbedUrl, t.Id));
                    appWorkSpacesTileList.SelectedIndex = 0;
                }
            }
            else
            {
                appWorkSpacesTileList.ItemsSource = null;
            }

        }


    }
}
