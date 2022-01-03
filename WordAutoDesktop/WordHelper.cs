using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutoDesktop
{
    class WordHelper
    {
        public WordHelper()
        {

        }

        private FileInfo fileInfo1;
        private FileInfo fileInfo2;
        public WordHelper(string fileName1, string fileName2)
        {
            if (File.Exists(fileName1))
            {
                fileInfo1 = new FileInfo(fileName1);
            }
            else
            {
                throw new ArgumentException("Main file not found!");
            }
            if (File.Exists(fileName2))
            {
                fileInfo2 = new FileInfo(fileName2);
            }
            else
            {
                throw new ArgumentException("Extra file not found!");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            Word.Document doc1 = null;

            try
            {
                app = new Word.Application();
                object file1 = fileInfo1.FullName;
                object file2 = fileInfo2.FullName;

                object missing = Type.Missing;
                doc1 = app.Documents.Open(file1, ReadOnly: true, Revert: true);

                foreach (var item in items)
                {
                    Word.Find find = doc1.Application.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: false,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                object newFileName = Path.Combine(fileInfo1.DirectoryName, "CREATED_" + fileInfo1.Name);
                doc1.Application.ActiveDocument.SaveAs2(newFileName);
                doc1.Application.ActiveDocument.Close();

                return true;
            }
            catch(Exception ex)
            { 
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }

            return false;
        }
        public bool Test()
        {
            Word.Application app = null;

            try
            {
                app = new Word.Application();
                object filePath = @"C:\Users\vlarikev\Desktop\WAD_Build\Extra File.docx";

                Word.Document doc = app.Documents.Open(filePath, ReadOnly: true, Visible: true);

                string keyWord1 = "KeyWord1";
                string keyWord2 = "KeyWord2";

                Word.Range range1 = doc.Content;
                range1.Find.Execute(keyWord1);

                Word.Range range2 = doc.Content;
                range2.Find.Execute(keyWord2);

                range1 = doc.Range(range1.End, range2.Start);
                range1.Text = " RANGE ";

                doc.SaveAs2(@"C:\Users\vlarikev\Desktop\WAD_Build\Extra File_CREATED.docx");
                doc.Close();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }

            return false;
        }
    }
}
