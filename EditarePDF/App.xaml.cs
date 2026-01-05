using System;
using System.Windows;
using Syncfusion.Licensing;

namespace EditarePDF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            // Register your Syncfusion license before any control is created.
            // Use an environment variable or user secrets to avoid hardcoding.
            var key = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
            if (!string.IsNullOrWhiteSpace(key))
            {
                SyncfusionLicenseProvider.RegisterLicense(key);
            }
        }
    }

}
