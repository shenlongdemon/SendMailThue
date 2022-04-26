using SendMailThue.Extentions;
using System;
using System.Threading;

namespace SendMailThue
{
    public class ErrorUtils
    {

        public static void ShowError(object ex,string info, params object []objs)
        {
            String s = $"{info} \n\nException={ex}\n\n";
            
            foreach (object obj in objs)
            {
                s += Newtonsoft.Json.JsonConvert.SerializeObject(obj) + "\n\n\n";
            }
            ShowError(s, true);
        }
        public static void ShowError(object ex, bool newThead)
        {
            ShowError($"ThreadExceptionEventHandler - Exception={ex}", newThead);
        }

         public static void ShowError(Exception ex, bool newThead)
        {

            ShowError(ex.GetTrace(), newThead);
        }
        public static void ShowError(Exception ex, string info, params object [] objs)
        {

            ShowError(ex.GetTrace(), info, objs);
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
