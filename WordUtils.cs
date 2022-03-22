using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;


namespace SendMailThue
{
    public class WordUtils
    {
        public static List<string> Replace(string wordFile, string outDir, List<List<string[]>> replaces)
        {
            List<string> outs = new List<string>();
            FileUtils.CreateDir(outDir, true);
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application();
                

                object missing = System.Reflection.Missing.Value;

                int outIndex = 0;
                foreach (List<string[]> replace in replaces)
                {

                    document = application.Documents.Add(wordFile);
                    foreach (string[] rep in replace)
                    {
                        Microsoft.Office.Interop.Word.Find findObject = application.Selection.Find;
                        findObject.ClearFormatting();
                        findObject.Text = rep[0];
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = rep[1];

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                    }
                    string fileName = outDir + @"\word_" + outIndex + ".doc";
                    document.SaveAs(fileName);
                    document.Close();
                    Marshal.ReleaseComObject(document);
                    outs.Add(fileName + "");
                    outIndex++;
                }
            }
            catch (Exception ex)
            {
            }
            finally {
                CloseWord(application, null);
            }
            return outs;
        }
        private static void UndoAllChange(Word.Document document)
        {
            object objTimes = 1;
            while (document.Undo(ref objTimes))
            {
            }
        }

        private static void CloseWord(Word.Application application, Word.Document document)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

          
            if (document != null)
            {
                document.Close();
                Marshal.ReleaseComObject(document);
            }


            //quit and release
            if (application != null)
            {
                application.Quit();
                Marshal.ReleaseComObject(application);
            }   
        }
    }
}
