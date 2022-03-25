﻿using System;
using System.Collections.Generic;
using System.Text;

namespace SendMailThue
{
    public class FileUtils
    {

        
        public static String ExecuteAppDir
        {
            get {
                return System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            }
        }
        public static String ExcelDir
        {
            get {
                return ExecuteAppDir + @"\excels";
            }
        } 
        public static String WordDir
        {
            get {
                return ExecuteAppDir + @"\words";
            }
        }
        public static void CreateDir(string outputDir, bool deleteIfExist)
        {
            bool exists = System.IO.Directory.Exists(outputDir);

            if (exists)
            {
                System.IO.Directory.Delete(outputDir, deleteIfExist);
            }
            System.IO.Directory.CreateDirectory(outputDir);
        }

        public static void CopyFile(string fromPathFile, string toPathFile)
        {
            System.IO.File.Copy(fromPathFile, toPathFile, true);
      }
        
    }
}
