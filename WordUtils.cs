using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;


namespace SendMailThue
{
    public class WordUtils
    {
        public static void Replace(string wordFile, string outDir, List<List<string[]>> replaces)
        {
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
                    string fileName = "";
                    foreach (string[] rep in replace)
                    {
                        if (rep[0] == "{fileName}")
                        {
                            fileName = rep[1];
                        }
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
                    string outFile = outDir + @"\" + fileName + ".doc";
                    document.SaveAs(outFile);
                    outIndex++;
                    CloseDocument(document);
                }
            }
            catch (Exception ex)
            {
            }
            finally {
                CloseWord(application, null);
            }
        }
        private static void UndoAllChange(Word.Document document)
        {
            object objTimes = 1;
            while (document.Undo(ref objTimes))
            {
            }
        }

        private static void CloseDocument(Word.Document document)
        {
            try
            {
                if (document != null)
                {
                    document.Close();
                    Marshal.ReleaseComObject(document);
                }
            }
            catch (Exception ex)
            { }
        }

        private static void CloseWord(Word.Application application, Word.Document document)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();


            CloseDocument(document);


            //quit and release
            try
            {
                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }
            }
            catch (Exception ex)
            { }
        }
    }
}
