using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
   
    class Test
    {
        private static String path = Directory.GetCurrentDirectory();
        private static Word.Application app = new Word.Application();
        private static Word.Document document = app.Documents.Open(path+"\\1.docx");
        private static Word.Document template = app.Documents.Open(path + "\\2.docx");

        public static void start_test()
        {
            try
            {
                font_test();
            }
            finally
            {
                document.Close();
                template.Close();
                app.Quit();
                Console.ReadKey();
            }

            
        }

        private static void font_test()
        {
            Font templateFont = new Font(template.Paragraphs[1].Range.Font);
            Word.Font font = document.Paragraphs[1].Range.Font;
            foreach (int i in templateFont.test_CompareTo(font))
            {
                Console.WriteLine(i);
            }
        }
    }
}
