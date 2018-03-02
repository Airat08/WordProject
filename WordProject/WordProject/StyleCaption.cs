using System;
using Microsoft.Office.Interop.Word;

namespace WordProject
{
    class StyleCaption : Style
    {
        private Template template;
        private Paragraph templateCaption;

        public StyleCaption()
        {
            template = Template.getInstance();
            templateCaption = template.getTemplateCaption();
        }

        public bool Equals(Paragraph paragraph)
        {
            Font font = new Font(templateCaption.Range.Font);
            ParagraphFormat paragraphFormat = new ParagraphFormat(templateCaption.Range.ParagraphFormat);
            if (font.CompareTo(paragraph.Range.Font) ||
                paragraphFormat.CompareTo(paragraph.Range.ParagraphFormat))
                return true;
            return false;
        }
    }
}
