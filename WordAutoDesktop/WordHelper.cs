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
        private FileInfo fileMainInfo;
        private FileInfo fileExtraInfo;
        public WordHelper(string fileMainName, string fileExtraName)
        {
            if (File.Exists(fileMainName))
            {
                fileMainInfo = new FileInfo(fileMainName);
            }
            else
            {
                throw new ArgumentException("Main file not found!");
            }
            if (File.Exists(fileExtraName))
            {
                fileExtraInfo = new FileInfo(fileExtraName);
            }
            else
            {
                throw new ArgumentException("Extra file not found!");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;

            try
            {
                app = new Word.Application();
                object fileMain = fileMainInfo.FullName;
                object fileExtra = fileExtraInfo.FullName;

                object missing = Type.Missing;

                string keyWord1 = "KeyWord1";
                string keyWord2 = "KeyWord2";

                string placeWord1 = "PlaceWord1";

                Word.Document docExtra = app.Documents.Open(fileExtra, ReadOnly: true, Visible: true);
                Word.Range range1key = docExtra.Content;
                range1key.Find.Execute(keyWord1);

                Word.Range range2key = docExtra.Content;
                range2key.Find.Execute(keyWord2);

                Word.Range rangeExtraPart = docExtra.Content;
                rangeExtraPart = docExtra.Range(range1key.End + 1, range2key.Start - 1);

                Word.Document docMain = app.Documents.Open(fileMain, ReadOnly: true, Visible: true);
                Word.Range rangePlaceWord = docMain.Content;

                rangePlaceWord.Find.Execute(placeWord1);
                rangePlaceWord = docMain.Range(rangePlaceWord.Start, rangePlaceWord.End);
                rangePlaceWord.Text = rangeExtraPart.Text;

                foreach (var item in items)
                {
                    Word.Find find = docMain.Application.Selection.Find;
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

                object newFileName = Path.Combine(fileMainInfo.DirectoryName, "CREATED_" + fileMainInfo.Name);
                docMain.SaveAs2(newFileName);
                docMain.Close();
                docExtra.Close();

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
    }
}
