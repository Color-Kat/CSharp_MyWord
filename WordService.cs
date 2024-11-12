using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace MyWord
{
    internal class WordService
    {
        // MS Word instance
        private Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

        // Word document instance
        private Document document;

        // private object missingValue = Type.Missing;
        private object endOfDocument = "\\endofdoc";

        public void CreateWordDocument(string paragraphText)
        {
            try
            {
                // Show MS Word application window
                wordApp.Visible = true;

                document = wordApp.Documents.Add();

                AddBlueHeading("1. Создание Word документа в C#", paragraphText);
                AddVioletHeading("1. Создание Excel документа в C#", paragraphText);

                AddTable(3, 3);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }



        public void AddBlueHeading(string heading, string text)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add();

            // Заголовок 1 - is a embeded in word paragraph style
            paragraph.Range.set_Style("Заголовок 1");

            // Configure paragraph
            paragraph.Range.Text = heading;
            paragraph.Range.Font.Bold = 1;
            paragraph.Format.SpaceAfter = 12;
            paragraph.Range.InsertParagraphAfter();

            if (text != "")
            {
                Paragraph textParagraph = document.Content.Paragraphs.Add();
                textParagraph.Range.Text = text;
                textParagraph.Format.SpaceAfter = 24;
                textParagraph.Range.InsertParagraphAfter();
            }
        }

        public void AddVioletHeading(string heading, string text)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add();

            // Заголовок 2 - is a embeded in word paragraph style
            paragraph.Range.set_Style("Заголовок 2");

            // Configure paragraph
            paragraph.Range.Text = heading;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Italic = 1;
            paragraph.Range.Font.ColorIndex = WdColorIndex.wdViolet;
            paragraph.Format.SpaceAfter = 12;
            paragraph.Range.InsertParagraphAfter();

            if (text != "")
            {
                Paragraph textParagraph = document.Content.Paragraphs.Add();
                textParagraph.Range.Text = text;
                textParagraph.Format.SpaceAfter = 24;
                textParagraph.Range.InsertParagraphAfter();
            }
        }

        public void AddTable(int rowCount, int colCount)
        {
            Range documentRange = document.Bookmarks.get_Item(ref endOfDocument).Range;

            Table table = document.Content.Tables.Add(documentRange, rowCount, colCount);
            table.Borders.Enable = 1;

            for (int i = 0; i < rowCount; i++)
            {
                for(int j = 0; j < colCount; j++)
                {
                    table.Range.Font.ColorIndex = WdColorIndex.wdBrightGreen;
                    table.Cell(i+1, j+1).Range.Text = "Row " + i + ", Col " + j;
                }
            }

        }

    }
}
