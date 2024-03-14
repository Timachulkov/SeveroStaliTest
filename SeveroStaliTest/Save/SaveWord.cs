using Microsoft.Office.Interop.Word;
using SeveroStaliTest.Filter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Application = Microsoft.Office.Interop.Word.Application;

namespace SeveroStaliTest
{
    class SaveWord : ISave
    {
        public bool Save(string DataSource, IEnumerable<FilteredData> data)
        {
            Saving(DataSource, data);
            return true;
        }
        private bool Saving(string DataSource, IEnumerable<FilteredData> data)
        {
            Application app = new Application();
            try
            {
                Document doc = app.Documents.Add();
                Paragraph para = doc.Content.Paragraphs.Add();

                para.Range.Text = "Отчет по загрузке";
                para.Range.Font.Size = 14;
                para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.InsertParagraphAfter();
                para.Range.Text = "";

                para.Range.InsertParagraphAfter();

                Table table;
                Range wrdRng = doc.Bookmarks.get_Item("\\endofdoc").Range;
                var Rows = data.Select(x => x.FilteredName.Count()).Sum(x => x) + data.Count();
                table = doc.Tables.Add(wrdRng, Rows + 1, 2);
                table.Borders.Enable = 1;
                table.Range.ParagraphFormat.SpaceAfter = 0;
                int c = 1;
                table.Cell(c, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                table.Cell(c, 1).Range.Font.Bold = 1;
                table.Cell(c, 1).Range.Font.Size = 11;
                table.Cell(c, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray50;
                table.Cell(c, 1).Range.Font.Color = WdColor.wdColorWhite;
                table.Cell(c, 1).Range.Text = "Отдел";
                table.Cell(c, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray50;
                table.Cell(c, 2).Range.Font.Color = WdColor.wdColorWhite;
                table.Cell(c, 2).Range.Font.Bold = 1;
                table.Cell(c, 2).Range.Font.Size = 11;
                table.Cell(c, 2).Range.Text = "Количество задач";

                foreach (var dep in data)
                {
                    c++;
                    table.Cell(c, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    table.Cell(c, 1).Range.Font.Bold = 1;
                    table.Cell(c, 1).Range.Font.Size = 11;
                    table.Cell(c, 1).Range.Text = dep.DepartmentsName;
                    table.Cell(c, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    table.Cell(c, 2).Range.Font.Bold = 1;
                    table.Cell(c, 2).Range.Font.Size = 11;
                    table.Cell(c, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    table.Cell(c, 2).Range.Text = dep.TaskNum.ToString();
                    foreach (var emp in dep.FilteredName)
                    {
                        c++;
                        table.Cell(c, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        table.Cell(c, 1).Range.Font.Size = 11;
                        table.Cell(c, 1).Range.Text = emp.Name;
                        table.Cell(c, 2).Range.Font.Size = 11;
                        table.Cell(c, 2).Range.Text = emp.TaskNum.ToString();
                    }
                }

                app.ActiveDocument.SaveAs2(DataSource);

                app.Quit(true);
                return true;
            }
            catch(Exception e)
            {
                app.Quit(false);
                MessageBox.Show(e.Message);
                return false;
            }
        }
    }
}
