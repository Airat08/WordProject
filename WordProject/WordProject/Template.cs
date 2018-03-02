using System;
using Microsoft.Office.Interop.Word;

namespace WordProject
{
    class Template
    {
        private static Template instance;
        private Document document;
        public Application wordApp;
        private const int TEMPLATE_CAPTION = 1;
        private const int TEMPLATE_TEXT = 2;
        private const int TEMPLATE_SUBSECTION = 3;
        private const int TEMPLATE_IMAGE = 4;
        private const int TEMPLATE_TABLE = 5;
        private const int TEMPLATE_PROGRAMMING_CODE = 6;
        //private const int TEMPLATE_SUBSECTION = 3;

        private Template()
        {
            wordApp = new Application();
            document = wordApp.Documents.Open(@"C:\Users\Айрат\Desktop\ТПО\WordProject\WordProject\WordProject\bin\Debug\2.docx");
        }

        public static Template getInstance()
        {
            if (instance == null)
                instance = new Template();
            return instance;
        }

        public Paragraph getTemplateCaption()
        {
            return document.Paragraphs[TEMPLATE_CAPTION];
        }

        public Paragraph getTemplateText()
        {
            return document.Paragraphs[TEMPLATE_TEXT];
        }

        public Paragraph getTemplateSubsection()
        {
            return document.Paragraphs[TEMPLATE_SUBSECTION];
        }

        public Paragraph getTemplateImage()
        {
            return document.Paragraphs[TEMPLATE_IMAGE];
        }

        public Paragraph getTemplateTable()
        {
            return document.Paragraphs[TEMPLATE_TABLE];
        }

        public Paragraph getTemplateProgrammingCode()
        {
            return document.Paragraphs[TEMPLATE_PROGRAMMING_CODE];
        }

        public Paragraphs getTemplateParagraphs()
        {
            return document.Paragraphs;
        }
    }
}
