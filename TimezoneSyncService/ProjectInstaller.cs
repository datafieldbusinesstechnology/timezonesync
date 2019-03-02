using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;

namespace SSTimezone
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {

            try
            {
                InitializeComponent();
                this.serviceInstaller1.AfterInstall += (sender, args) => new System.ServiceProcess.ServiceController(serviceInstaller1.ServiceName).Start();

                System.ServiceProcess.ServiceController[] sconarray = System.ServiceProcess.ServiceController.GetServices();
                for (int i = 0; i < sconarray.Length - 1; i++)
                {
                    if (sconarray[i].ServiceName == serviceInstaller1.ServiceName)
                    {
                        if (sconarray[i].Status == System.ServiceProcess.ServiceControllerStatus.Running)
                        {
                            this.serviceInstaller1.BeforeUninstall += (sender, args) => sconarray[i].Stop();
                        }
                    }
                }

            }
            catch (Exception e) {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
    }
}