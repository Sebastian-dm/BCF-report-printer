using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Reflection;
using System.IO;
using System.Globalization;

namespace BCFReader
{

    //public class Program
    //{

    //    [STAThreadAttribute]
    //    public static void Main()
    //    {
    //        App.Main();
    //    }

    //    private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
    //    {
    //        Assembly executingAssembly = Assembly.GetExecutingAssembly();
    //        AssemblyName assemblyName = new AssemblyName(args.Name);

    //        string path = assemblyName.Name + ".dll";

    //        if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
    //        {

    //            path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);
    //        }
    //        using (Stream stream = executingAssembly.GetManifestResourceStream(path))
    //        {
    //            if (stream == null) return null;

    //            byte[] assemblyRawBytes = new byte[stream.Length];
    //            stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);

    //            return Assembly.Load(assemblyRawBytes);
    //        }
    //    }
    //}

    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void OnStartup(object sender, StartupEventArgs e)
        {
            //AppDomain.CurrentDomain.AssemblyResolve +=
            //    new ResolveEventHandler(ResolveAssembly);

            MainWindow main = new MainWindow();
            main.ShowDialog();
        }

        //static Assembly ResolveAssembly(object sender, ResolveEventArgs args)
        //{
        //    Assembly parentAssembly = Assembly.GetExecutingAssembly();

        //    var name = args.Name.Substring(0, args.Name.IndexOf(',')) + ".dll";
        //    var resourceName = parentAssembly.GetManifestResourceNames()
        //        .First(s => s.EndsWith(name));

        //    using (Stream stream = parentAssembly.GetManifestResourceStream(resourceName))
        //    {
        //        byte[] block = new byte[stream.Length];
        //        stream.Read(block, 0, block.Length);
        //        return Assembly.Load(block);
        //    }
        //}
    }
}
