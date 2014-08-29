namespace Ops2Zenoss
{
    using Microsoft.EnterpriseManagement;
    using Microsoft.EnterpriseManagement.Common;
    using Microsoft.EnterpriseManagement.ConnectorFramework;
    using mod_zenoss;
    using System;
    using System.Collections.ObjectModel;
    using System.Configuration;
    using System.Diagnostics;
    using System.Net;
    using System.ServiceProcess;
    using System.Threading;
    public partial class Ops2Zenoss : ServiceBase
    {
        private int pollingIntervalMS = 60000;
        private int startupDelayMS = 300000;

        // Constants for event ids
        private const int EventCheckAlertsRunningLong = 1001;
        private const int EventCheckAlertsMonitoringException = 1002;
        private const int EventOnStartInitilization = 1003;
        private const int EventConnectorNotInitialized = 1004;
        private const int EventInfoAboutToGetAlerts = 1005;
        private const int EventBadConfig = 1006;

        public static Guid connectorGuid
        {
            get { return new Guid("{CCD58864-A2E6-45af-8FE4-9C2EF78CD609}"); }
        }

        public Ops2Zenoss()
        {
            InitializeComponent();
            eventLog1.Source = "Ops2ZenossSource";
            try
            {
                pollingIntervalMS = int.Parse(ConfigurationManager.AppSettings["PollingIntervalMS"]);
                startupDelayMS = int.Parse(ConfigurationManager.AppSettings["StartupDelayMS"]);
            }
            catch (Exception excp)
            {
                eventLog1.WriteEntry("Missing or badly formed config file settings: " + excp.Message, 
                    EventLogEntryType.Error, 
                    EventBadConfig);
            }
        }

        private Timer chkAlertsTimer = null;

        MonitoringConnector connector;
        protected override void OnStart(string[] args)
        {
            eventLog1.Source = "Ops2ZenossSource";
            eventLog1.WriteEntry("Ops2Zenoss Connector Service Startup", EventLogEntryType.Information);
            TimerCallback runChkAlertsDelegate = new TimerCallback(CheckAlerts);

            chkAlertsTimer = new Timer(runChkAlertsDelegate, null, startupDelayMS, pollingIntervalMS);
        }

        protected override void OnStop()
        {
            eventLog1.Source = "Ops2ZenossSource";
            eventLog1.WriteEntry("Ops2Zenoss Connector Service Shutdown", EventLogEntryType.Information);
        }
        private void CheckAlerts(object data)
        {
            eventLog1.Source = "Ops2ZenossSource";
            if (!(Monitor.TryEnter(chkAlertsTimer)))
            {
                eventLog1.WriteEntry("CheckAlerts is running too long. Timer function found that prior execution is still active.",
                    EventLogEntryType.Error,
                    EventCheckAlertsRunningLong);
                return;
            }

            try
            {
                if (connector == null)
                {
                    ManagementGroup mg = new ManagementGroup("localhost");
                    IConnectorFrameworkManagement icfm = mg.ConnectorFramework;

                    connector = icfm.GetConnector(connectorGuid);
                }

                ReadOnlyCollection<ConnectorMonitoringAlert> Alerts;
                Alerts = connector.GetMonitoringAlerts();

                if (Alerts.Count > 0)
                {
                    eventLog1.WriteEntry("Found " + Alerts.Count + " alerts for processing", 
                        EventLogEntryType.Information);
                    OutputAlerts(Alerts);
                    eventLog1.WriteEntry("Acknowledging all alerts.", 
                        EventLogEntryType.Information);
                    connector.AcknowledgeMonitoringAlerts(Alerts);
                }
            }
            catch (EnterpriseManagementException error)
            {
                connector = null;
                eventLog1.WriteEntry("MonitoringException in timer function:" + error.Message,
                    EventLogEntryType.Error,
                    EventCheckAlertsMonitoringException);
            }
            catch (Exception excp)
            {
                connector = null;
                eventLog1.WriteEntry("Exception in timer function:" + excp.Message,
                    EventLogEntryType.Error, 
                    EventCheckAlertsMonitoringException);
            }
            finally
            {
                Monitor.Exit(chkAlertsTimer);
            }
        }
        private void OutputAlerts(ReadOnlyCollection<ConnectorMonitoringAlert> Alerts)
        {
            eventLog1.Source = "Ops2ZenossSource";
            try
            {
                NetworkCredential zenCreds = new NetworkCredential(ConfigurationManager.AppSettings["zenUser"], ConfigurationManager.AppSettings["zenPass"]);
                ZenossAPI.Connect(zenCreds, ConfigurationManager.AppSettings["zenServer"]);

                foreach (ConnectorMonitoringAlert Alert in Alerts)
                {
                    eventLog1.WriteEntry("Sending alert to Zenoss\r\n" + Alert.Description,
                        EventLogEntryType.Information);
                    ZenossAPI.CreateEvent(Alert.NetbiosComputerName, Alert.Severity.ToString(), Alert.Description);
                    Alert.Update("Alert has been forwarded to Zenoss");
                }
            }
            catch (Exception excp)
            {
                eventLog1.WriteEntry("Exception while attempting to send alerts to Zenoss:" + excp.Message,
                    EventLogEntryType.Error, 
                    EventCheckAlertsMonitoringException);
            }
        }
    }
}