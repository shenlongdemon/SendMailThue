using System;
using System.Collections.Generic;
using System.Text;

namespace SendMailThue.Extentions
{
    public static class ExceptionEx
    {
        public static string GetTrace(this Exception ex)
        {
            var stringBuilder = new StringBuilder();

            while (ex != null)
            {
                stringBuilder.AppendLine(ex.Message);
                stringBuilder.AppendLine(ex.StackTrace);
                stringBuilder.AppendLine(ex.ToString());
                ex = ex.InnerException;
            }

            return stringBuilder.ToString();
        }
    }
}
