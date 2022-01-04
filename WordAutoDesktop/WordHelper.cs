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

        // Key and place words.
        private string[] keyWordStartArray = new string[] { "KeyWord1", "Таблица 1. Первая таблица." };
        private string[] keyWordEndArray = new string[] { "KeyWord2", "Уникальный текст идет далее." };

        private string[] placeWordsArray = new string[] { "PlaceWord1", "PlaceWord2" };


        // Table vatiables.
        private string tableFontName = "Times New Roman";
        private int tableFontSize = 10;

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

                Word.Document docExtra = app.Documents.Open(fileExtra, ReadOnly: true, Visible: true);
                Word.Document docMain = app.Documents.Open(fileMain, ReadOnly: true, Visible: true);

                FindWordAndPasteToIt(docMain, placeWordsArray[0], FindPartBetweenKeywords(docExtra, keyWordStartArray[0], keyWordEndArray[0]));
                FindWordAndPasteToIt(docMain, placeWordsArray[1], FindPartBetweenKeywords(docExtra, keyWordStartArray[1], keyWordEndArray[1]));

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

        private Word.Range FindWordAndPasteToIt(Word.Document docMain, string placeWord, Word.Range rangeExtraPart)
        {
            Word.Range rangePlaceWord = docMain.Content;

            rangePlaceWord.Find.Execute(placeWord);
            rangePlaceWord = docMain.Range(rangePlaceWord.Start, rangePlaceWord.End);
            
            if (rangeExtraPart.Tables.Count == 0)
            {
                rangePlaceWord.Text = rangeExtraPart.Text;
            }
            else
            {
                rangeExtraPart.Copy();
                Word.Table table = rangePlaceWord.Tables.Add(rangePlaceWord, 1, 1);
                table.Range.Paste();

                for (int i = 0; i < table.Rows.Count + 1; i++)
                {
                    for (int j = 0; j < table.Columns.Count + 1; j++)
                    {
                        Word.Cell cell = table.Cell(i, j);
                        cell.Range.Font.Name = tableFontName;
                        cell.Range.Font.Size = tableFontSize;
                        cell.Range.Font.Bold = 0;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    }
                }
            }

            return rangePlaceWord;
        }
        private Word.Range FindPartBetweenKeywords(Word.Document docExtra, string keyStart, string keyEnd)
        {
            Word.Range range1key = docExtra.Content;
            range1key.Find.Execute(keyStart);

            Word.Range range2key = docExtra.Content;
            range2key.Find.Execute(keyEnd);

            Word.Range rangeExtraPart = docExtra.Range(range1key.End + 1, range2key.Start - 1);

            return rangeExtraPart;
        }
    }
}
