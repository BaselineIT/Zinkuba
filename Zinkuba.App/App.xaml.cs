using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        // Exchange Webservices needs to be a physical dll, because a part of it checks for its existence, if you load it as a stream it af-kaks
        static private readonly String[] PhysicalAssemblies = { "Microsoft.Exchange.WebServices" };

        private void App_OnStartup(object sender, StartupEventArgs e)
        {
            // Make sure we load certain namespaces as resources (they are embedded dll's)
            AppDomain.CurrentDomain.AssemblyResolve += EmbeddedAssemblyResolver;
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
                        if (Array.Exists(PhysicalAssemblies, element => element.Equals(assemblyName.Name)))
                        {
                            String tempFile = Path.GetTempFileName();
                            File.WriteAllBytes(tempFile, assemblyData);
                            Console.WriteLine("[" + Thread.CurrentThread.ManagedThreadId + "-" +
                                              Thread.CurrentThread.Name + "] Loading assembly " + assemblyName.Name +
                                              " from " + tempFile);
                            return Assembly.LoadFile(tempFile);
                        }
                        else
                        {
                            return Assembly.Load(assemblyData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + (ex.InnerException != null ? " : " + ex.InnerException.Message : ""), "Program Failed to access Assembly " + args.Name, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return null;
        }
    }
}
