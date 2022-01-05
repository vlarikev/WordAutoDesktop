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
        private string[] keyWordStartArray = new string[] { "KeyWord1", "Таблица 1. Первая таблица.", "Далее идет рисунок." };
        private string[] keyWordEndArray = new string[] { "KeyWord2", "Уникальный текст идет далее.", "Рисунок 1. Тестовый рисунок." };

        private string[] placeWordsArray = new string[] { "PlaceWord1", "PlaceWord2", "PlaceWord3" };

        // Text variables.
        private string textFontName = "Times New Roman";
        private int textFontSize = 14;

        // Table variables.
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
                FindWordAndPasteToIt(docMain, placeWordsArray[2], FindPartBetweenKeywords(docExtra, keyWordStartArray[2], keyWordEndArray[2]));

                foreach (var item in items)
                {
                    Word.Find find = docMain.Application.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;
                    TextStylization(find.Replacement.Font, find.Replacement.ParagraphFormat, false);

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
                        Format: true,
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
            
            if (rangeExtraPart.Tables.Count > 0)
            {
                rangeExtraPart.Copy();
                Word.Table table = rangePlaceWord.Tables.Add(rangePlaceWord, 1, 1);
                table.Range.Paste();

                for (int i = 0; i < table.Columns.Count + 1; i++)
                {
                    for (int j = 0; j < table.Rows.Count + 1; j++)
                    {
                        Word.Cell cell = table.Cell(j, i);

                        cell.Range.Font.Bold = 0;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        TextStylization(cell.Range.Font, cell.Range.ParagraphFormat, true);
                    }

                    Word.Cell cellRow = table.Cell(1, i);
                    cellRow.Range.Font.Bold = 1;
                }
            }
            else if (rangeExtraPart.InlineShapes.Count > 0)
            {
                rangeExtraPart.Copy();
                rangePlaceWord.Paste();
                rangePlaceWord.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            else
            {
                rangePlaceWord.Text = rangeExtraPart.Text;
                TextStylization(rangePlaceWord.Font, rangePlaceWord.ParagraphFormat, false);
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
        private void TextStylization(Word.Font font, Word.ParagraphFormat paragraphFormat, bool isTable)
        {
            if (isTable)
            {
                font.Name = tableFontName;
                font.Size = tableFontSize;
            }
            else
            {
                font.Name = textFontName;
                font.Size = textFontSize;
            }

            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            paragraphFormat.LeftIndent = 0;
            paragraphFormat.RightIndent = 0;
            paragraphFormat.FirstLineIndent = 35.5f;
            paragraphFormat.SpaceBefore = 0;
            paragraphFormat.SpaceAfter = 0;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
        }
    }
}
