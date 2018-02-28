using Microsoft.Office.Interop.Word;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
    


    class Program
    {
        public Application wordApp;
        public Document doc;
        public Document model;

        public Program()
        {
            wordApp = new Application();
            doc = wordApp.Documents.Open(@"C:\Users\Айрат\Desktop\ТПО\WordProject\WordProject\WordProject\bin\Debug\2.docx");
            model = wordApp.Documents.Open(@"C:\Users\Айрат\Desktop\ТПО\WordProject\WordProject\WordProject\bin\Debug\model.docx");
            wordApp.Visible = true;
        }
        static void Main(string[] args)
        {
            ParagraphFormat paragraphFormat;
            Program program = new Program();
            paragraphFormat = new ParagraphFormat(program.model.Paragraphs[1].Range.ParagraphFormat);

            //Console.WriteLine(program.model.Paragraphs[1].Range.ParagraphFormat.HalfWidthPunctuationOnTopOfLine);
            //program.doc.Paragraphs[4].Range.ParagraphFormat.CharacterUnitLeftIndent = 20;
            //Console.WriteLine(program.doc.Paragraphs[3].Range.ParagraphFormat.HalfWidthPunctuationOnTopOfLine);

            Console.WriteLine(paragraphFormat.CompareTo(program.doc.Paragraphs[3].Range.ParagraphFormat));

            //Console.WriteLine(program.doc.Paragraphs[4].Range.Text);

            // Console.WriteLine(program.doc.Paragraphs[3].Range.ParagraphFormat.Alignment.CompareTo(
            //   program.model.Paragraphs[1].Range.ParagraphFormat.Alignment));
            //Console.WriteLine(program.doc.Paragraphs[25].Range.ParagraphFormat.Borders.Count);
            //Console.WriteLine(program.model.Paragraphs[1].Range.ParagraphFormat.Borders.Count);
            Console.Read();
        }
    }
}
