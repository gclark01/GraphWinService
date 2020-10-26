using System;
using System.Globalization;
using System.IO;
using System.ServiceProcess;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace GraphWinService
{
    public partial class Service1 : ServiceBase
    {

        public Service1()
        {
            InitializeComponent();
            this.ServiceName = "GraphWinService.NET";
        }

        protected override void OnStart(string[] args)
        {
            WriteLog("Service has been started");

            try
            {
                var _graphClient = GraphClientApp();
                getGroupsAsync(_graphClient).GetAwaiter();
            }
            catch (Exception e)
            {

                WriteLog(e.Message.ToString());
            }
            

        }

        public GraphServiceClient GraphClientApp()
        {
            string scopes = "https://graph.microsoft.com/.default";
            string clientId = "<<clientId>>";
            string tenantId = "<<tenantId>>";
            string secret = "<<secret>>";

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(secret)
                .Build();

            ClientCredentialProvider clientCredentialProvider = new ClientCredentialProvider(confidentialClientApplication, scopes);
            GraphServiceClient graphServiceClient = new GraphServiceClient(clientCredentialProvider);

            return graphServiceClient;

        }

        public async Task getGroupsAsync(GraphServiceClient graphServiceClient)
        {
            var groups = await graphServiceClient.Groups.Request().Select(x => new { x.Id, x.DisplayName }).GetAsync();
            foreach (var group in groups)
            {
                WriteLog($"{group.DisplayName}, {group.Id}");
            }
        }

        public void WriteLog(string logMessage, bool addTimeStamp = true)
        {
            var path = AppDomain.CurrentDomain.BaseDirectory;
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);

            var filePath = String.Format("{0}\\{1}_{2}.txt",
                path,
                ServiceName,
                DateTime.Now.ToString("yyyyMMdd", CultureInfo.CurrentCulture)
                );

            if (addTimeStamp)
                logMessage = String.Format("[{0}] - {1}",
                    DateTime.Now.ToString("HH:mm:ss", CultureInfo.CurrentCulture),
                    logMessage);

            System.IO.File.AppendAllText(filePath, logMessage + Environment.NewLine);
        }

        protected override void OnStop()
        {
            WriteLog("Service has been stopped");
        }
    }
}
