using System;
using System.Collections.Generic;
using System.Text;

namespace SendMailThue
{
    public class FileUtils
    {
        public static void CreateDir(string outputDir, bool deleteIfExist)
        {
            bool exists = System.IO.Directory.Exists(outputDir);

            if (exists)
            {
                System.IO.Directory.Delete(outputDir, deleteIfExist);
            }
            System.IO.Directory.CreateDirectory(outputDir);
        }
    }
}
