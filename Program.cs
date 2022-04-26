using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Permissions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendMailThue
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        /// 
        [STAThread]
        //[HandleProcessCorruptedStateExceptions]
        //[SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.ControlAppDomain)]
        static void Main()
        {

            Application.ThreadException += new ThreadExceptionEventHandler(ThreadExceptionEventHandler);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            AppDomain.CurrentDomain.FirstChanceException += FirstChanceExceptionEventHandler;
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledExceptionEventHandler);

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.Run(new SendMailForm());
            }
            catch (Exception ex) {
                new ErrorBox($"UnhandledExceptionEventHandler - Exception={ex}").Show();
            }
        }

        private static void UnhandledExceptionEventHandler(object sender, UnhandledExceptionEventArgs e)
        {
            //ErrorUtils.ShowError(e.ExceptionObject, true);
            //new ErrorBox($"UnhandledExceptionEventHandler - Exception={e.ExceptionObject}").Show();
            //Console.WriteLine($"UnhandledExceptionEventHandler - Exception={e.ExceptionObject}");
        }

        private static void FirstChanceExceptionEventHandler(object sender, FirstChanceExceptionEventArgs e)
        {
            //ErrorUtils.ShowError(e.Exception, true);
            //new ErrorBox($"FirstChanceExceptionEventHandler - Exception={e.Exception}").Show();
            //Console.WriteLine($"FirstChanceExceptionEventHandler - Exception={e.Exception}");
        }

        private static void ThreadExceptionEventHandler(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            //ErrorUtils.ShowError(e.Exception, true);
            //new ErrorBox($"ThreadExceptionEventHandler - Exception={e.Exception}").Show();
            //MessageBox.Show($"ThreadExceptionEventHandler - Exception={e.Exception}");
        }
    }
}
