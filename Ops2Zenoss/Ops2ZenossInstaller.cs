namespace Ops2Zenoss
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Configuration.Install;
    using Microsoft.EnterpriseManagement;
    using Microsoft.EnterpriseManagement.ConnectorFramework;
    using Microsoft.EnterpriseManagement.Common;
    using Microsoft.EnterpriseManagement.Monitoring;
    using System.ServiceProcess;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    [RunInstaller(true)]
    public partial class Ops2ZenossInstaller : System.Configuration.Install.Installer
    {
        private ServiceInstaller serviceInstaller;
        private ServiceProcessInstaller processInstaller;

        public Ops2ZenossInstaller()
        {
            InitializeComponent();
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();

            // Service will be setup to run as local system. You will
            // probably need to change this (via service manager console) to something that has rights to
            // access Operations Manager.
            processInstaller.Account = ServiceAccount.LocalSystem;

            // Service will have Start Type of Automatic
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            serviceInstaller.ServiceName = "Ops2Zenoss";
            serviceInstaller.Description = "System Center Operations Manager 2012 Connector for Zenoss";
            serviceInstaller.DisplayName = "Send Operations Manager alerts to Zenoss";

            Installers.Add(serviceInstaller);
            Installers.Add(processInstaller);
        }
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            ManagementGroup mg = new ManagementGroup("localhost");
            IConnectorFrameworkManagement icfm = mg.ConnectorFramework;
            MonitoringConnector connector;

            ConnectorInfo info = new ConnectorInfo();
            info.Description = serviceInstaller.Description;
            info.DisplayName = serviceInstaller.DisplayName;
            info.Name = serviceInstaller.ServiceName;

            connector = icfm.Setup(info, Ops2Zenoss.connectorGuid);
            connector.Initialize();

            eventLog1 = new EventLog();
            if (!(EventLog.SourceExists("Ops2ZenossSource")))
            {
                EventLog.CreateEventSource("Ops2ZenossSource", "Application");
            }
            eventLog1.Source = "Ops2ZenossSource";
            eventLog1.Log = "Application";

            eventLog1.WriteEntry("Installing Connector " + connector.Name,
                EventLogEntryType.Information);
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            eventLog1.Source = "Ops2ZenossSource";
            ManagementGroup mg = new ManagementGroup("localhost");
            IConnectorFrameworkManagement icfm = mg.ConnectorFramework;
            MonitoringConnector connector;

            connector = icfm.GetConnector(Ops2Zenoss.connectorGuid);
            IList<MonitoringConnectorSubscription> Subscriptions;

            eventLog1.WriteEntry("Uninstalling Connector : " + connector.Name,
                EventLogEntryType.Information);

            Subscriptions = icfm.GetConnectorSubscriptions();
            foreach (MonitoringConnectorSubscription Subscription in Subscriptions)
            {
                if (Subscription.MonitoringConnectorId == Ops2Zenoss.connectorGuid)
                {
                    icfm.DeleteConnectorSubscription(Subscription);
                    eventLog1.WriteEntry("Removing Subscription : " + Subscription.DisplayName,
                        EventLogEntryType.Information);
                }
            }
            connector.Uninitialize();
            icfm.Cleanup(connector);
            eventLog1.WriteEntry("Connector Removed",
                EventLogEntryType.Information);
        }
    }
}
