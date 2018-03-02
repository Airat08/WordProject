using System;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
    class Navigation
    {
        private int nextParagraphNumber = 1;
        public Application wordApp;
        public Document document;
        public Paragraphs paragraphs;

        public Navigation()
        {
            wordApp = new Application();
            document = wordApp.Documents.Open(@"C:\Users\Айрат\Desktop\ТПО\WordProject\WordProject\WordProject\bin\Debug\2.docx");
            wordApp.Visible = true;

        }

        public void start()
        {
            wordApp.Quit();
            Word.Font font = new Word.Font();
            Word.Paragraph paragraph;
            while (hasNext())
            {
                paragraph = next();
                if (isText(paragraph)) continue;
                if (isCaption(paragraph)) continue;
                if (isFormula(paragraph)) continue;
                if (isCode(paragraph)) continue;
                //addException(paragraph);
            };
        }

        private bool hasNext()
        {
            if (document.Paragraphs.Count <= nextParagraphNumber)
                return true;
            return false;
        }

        private Word.Paragraph next()
        {
            return document.Paragraphs[nextParagraphNumber++];
        }

        private bool isCode(Paragraph paragraph)
        {
            return false;
        }

        private bool isFormula(Paragraph paragraph)
        {
            return false;
        }

        private bool isCaption(Paragraph paragraph)
        {
            return false;
        }

        private bool isText(Paragraph paragraph)
        {
            return false;
        }

        private bool isTable(Paragraph paragraph)
        {
            if (paragraph.Range.Tables.Count == 0)
                return false;
            return true;

        }

        private bool isImage(Paragraph paragraph)
        {
            if (paragraph.Range.InlineShapes.Count != 0)
                return true;
            return false;
        }

    }
}
