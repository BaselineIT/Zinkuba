using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookTester
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public void Init(object sender, StartupEventArgs startupEventArgs)
        {
            // Make sure we load certain namespaces as resources (they are embedded dll's)
            AppDomain.CurrentDomain.AssemblyResolve += EmbeddedAssemblyResolver;
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }

        private static Assembly EmbeddedAssemblyResolver(object sender, ResolveEventArgs args)
        {
            try
            {
                var assemblyName = new AssemblyName(args.Name);
                String resourceName = Assembly.GetExecutingAssembly().FullName.Split(',').First() + "." + assemblyName.Name + ".dll";
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        Byte[] assemblyData = new Byte[stream.Length];
                        stream.Read(assemblyData, 0, assemblyData.Length);
                            return Assembly.Load(assemblyData);
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return null;
        }
    }
}
