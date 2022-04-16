using SendMailThue.Extentions;
using System;
using System.Threading;

namespace SendMailThue
{
    public class ErrorUtils
    {

        public static void ShowError(object ex, bool newThead)
        {
            ShowError($"ThreadExceptionEventHandler - Exception={ex}", newThead);
        }

         public static void ShowError(Exception ex, bool newThead)
        {

            ShowError(ex.GetTrace(), newThead);
        }
        
        public static void ShowError(string msg, bool newThead)
        {
            if (newThead)
            {
                new Thread(() => new ErrorBox(msg).ShowDialog()).Start();
            }
            else
            {
                new ErrorBox(msg).ShowDialog();
            }
        }
    }
}
