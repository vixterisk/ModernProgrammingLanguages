using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;

namespace WindowsFormsApp1
{
    public static class Generator
    {
        static string message;
        public static bool UniquePossibility { get; private set; }
        static List<string> names;
        static List<List<string>> wishes;
        static int highestWishCount = 0;
        static string font;

        static string ReadFromExcel(string excelFileName)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            string wordFileName = "";
            try
            {
                wishes = new List<List<string>>();
                names = new List<string>();
                var excelDoc = excelApp.Workbooks.Open(excelFileName);
                var namesWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Sheets["Имена"];
                var wishesWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Sheets["Группы"];
                var settingsWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Sheets["Настройки"];
                var namesLastCell = namesWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                int lastRow = namesLastCell.Row;
                Microsoft.Office.Interop.Excel.Range cells;
                for (int i = 0; i < lastRow; i++)
                {
                    cells = (Microsoft.Office.Interop.Excel.Range)namesWorksheet.Cells[i + 1, 1];
                    if (cells.Text != "")
                        names.Add((string)cells.Text);
                }
                var wishesLastCell = wishesWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                lastRow = wishesLastCell.Row;
                int lastColumn = wishesLastCell.Column;
                for (int i = 0; i < lastColumn; i++)
                {
                    wishes.Add(new List<string>());
                    for (int j = 1; j < lastRow; j++)
                    {
                        cells = (Microsoft.Office.Interop.Excel.Range)wishesWorksheet.Cells[j + 1, i + 1];
                        if (cells.Text != "")
                        {
                            wishes[i].Add((string)cells.Text);
                            if (highestWishCount < wishes[i].Count) highestWishCount = wishes[i].Count;
                        }
                    }
                }
                cells = (Microsoft.Office.Interop.Excel.Range)settingsWorksheet.Cells[2, 1];
                wordFileName = (string)cells.Text;
                cells = (Microsoft.Office.Interop.Excel.Range)settingsWorksheet.Cells[2, 2];
                font = (string)cells.Text;
                int result = 1;
                for (int i = 0; i < wishes.Count; i++)
                    result *= wishes[i].Count;
                UniquePossibility = result >= names.Count;
                excelDoc.Close(0);
            }
            catch (Exception e) {
                message = "Упс! Что-то пошло не так...";
            }
            excelApp.Quit();
            return wordFileName;
        }

        static void RefillGroup(List<List<string>> unusedWishes, int groupNumber)
        {
            if (unusedWishes[groupNumber].Count == 0)
                foreach (var wish in wishes[groupNumber])
                    unusedWishes[groupNumber].Add(wish);
        }

        static void RefillListOfGroups(List<int> unusedGroups)
        {
            if (unusedGroups.Count < 3)
                for (int i = 0; i < wishes.Count; i++)
                    unusedGroups.Add(i);
        }
        static void RefillGroupsIfNecessary(List<List<string>> unusedWishes, int firstTheme, int secondTheme, int thirdTheme)
        {
            RefillGroup(unusedWishes, firstTheme);
            RefillGroup(unusedWishes, secondTheme);
            RefillGroup(unusedWishes, thirdTheme);
        }

        static List<List<string>> GeneratedCongratilations()
        {
            var rand = new Random();
            var unusedGroups = new List<int>();
            var unusedWishes = new List<List<string>>();
            for (int i = 0; i < wishes.Count; i++)
            {
                unusedGroups.Add(i);
                var newGroup = new List<string>();
                unusedWishes.Add(newGroup);
                foreach (var wish in wishes[i])
                    newGroup.Add(wish);
            }
            var result = new List<List<string>>();
            foreach (var name in names)
            {
                var curList = new List<string>();
                curList.Add(name);
                result.Add(curList);
            }
            int firstThemeNumber = 0, secondThemeNumber = 1, thirdThemeNumber = 2, firstTheme = 0, secondTheme = 1, thirdTheme = 2;
            int firstRandomWish = 0, secondRandomWish = 1, thirdRandomWish = 2;
            for (int i = 0; i < result.Count; i++)
            {
                RefillListOfGroups(unusedGroups);
                firstThemeNumber = rand.Next(unusedGroups.Count);
                firstTheme = unusedGroups[firstThemeNumber];
                unusedGroups.RemoveAt(firstThemeNumber);
                secondThemeNumber = rand.Next(unusedGroups.Count);
                secondTheme = unusedGroups[secondThemeNumber];
                unusedGroups.RemoveAt(secondThemeNumber);
                thirdThemeNumber = rand.Next(unusedGroups.Count);
                thirdTheme = unusedGroups[thirdThemeNumber];
                unusedGroups.RemoveAt(thirdThemeNumber);
                RefillGroupsIfNecessary(unusedWishes, firstTheme, secondTheme, thirdTheme);
                var maxInFirst = unusedWishes[firstTheme].Count;
                firstRandomWish = rand.Next(maxInFirst);
                var firstWish = unusedWishes[firstTheme][firstRandomWish];
                unusedWishes[firstTheme].Remove(firstWish);
                result[i].Add(firstWish);
                var maxInSecond = unusedWishes[secondTheme].Count;
                secondRandomWish = rand.Next(maxInSecond);
                var secondWish = unusedWishes[secondTheme][secondRandomWish];
                unusedWishes[secondTheme].Remove(secondWish);
                result[i].Add(secondWish);
                var maxInThird = unusedWishes[thirdTheme].Count;
                thirdRandomWish = rand.Next(maxInThird);
                var thirdWish = unusedWishes[thirdTheme][thirdRandomWish];
                unusedWishes[thirdTheme].Remove(thirdWish);
                result[i].Add(thirdWish);
            }                 
            return result;    
        }

        static void WriteInWord(string wordFileName, List<List<string>> generatedCongratilations)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Add(wordFileName);
                for (int i = 0; i < generatedCongratilations.Count; i++)
                {
                    var bookmark = wordApp.ActiveDocument.Bookmarks["Name"];
                    bookmark.Range.Text = generatedCongratilations[i][0];
                    bookmark = wordApp.ActiveDocument.Bookmarks["Wish1"];
                    bookmark.Range.Text = generatedCongratilations[i][1];
                    bookmark = wordApp.ActiveDocument.Bookmarks["Wish2"];
                    bookmark.Range.Text = generatedCongratilations[i][2];
                    bookmark = wordApp.ActiveDocument.Bookmarks["Wish3"];
                    bookmark.Range.Text = generatedCongratilations[i][3];
                    if (i != generatedCongratilations.Count - 1)
                    {
                        wordApp.Selection.EndKey(WdUnits.wdStory);
                        wordApp.Selection.InsertNewPage();
                        wordApp.Selection.InsertFile(wordFileName, "", true, false, true);
                    }
                }
                for (int i = 0; i < wordDoc.Paragraphs.Count; i++)
                {
                    Microsoft.Office.Interop.Word.Range rng = wordDoc.Paragraphs[i+1].Range;
                    rng.Font.Name = font;
                }
                //wordApp.ActiveWindow.View.Type = WdViewType.wdWebView;
                //wordApp.Visible = true;
                string currentPath = Directory.GetCurrentDirectory();
                var path = Path.Combine(currentPath, "Congratilations");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                string fileName = DateTime.Now.ToString("yyyy.MM.dd HH.mm.ss");
                wordDoc.SaveAs(path + "//" + fileName + ".docx");
                wordDoc.Close();
                message = "Поздравления сгенерированы!";
            }
            catch (Exception e)
            {
                message = "Упс! Что-то пошло не так...";
                wordApp.Quit();
            }
        }

        public static string CreateCongratilation(string excelFileName)
        {
            message = "";
            string wordFileName = ReadFromExcel(excelFileName);
            var generatedCongratilations = GeneratedCongratilations();
            WriteInWord(wordFileName, generatedCongratilations);
            return message;
        }
    }
}
