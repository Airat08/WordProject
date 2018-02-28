using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
    class ParagraphFormat
    {
        private Word.ParagraphFormat defaultParagraphFormat;

        public ParagraphFormat(Word.ParagraphFormat paragraphFormat)
        {
            defaultParagraphFormat = paragraphFormat;
        }

        public bool CompareTo(Word.ParagraphFormat paragraphFormat)
        {
            return (paragraphFormat.AddSpaceBetweenFarEastAndAlpha.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndAlpha) &&
                    paragraphFormat.AddSpaceBetweenFarEastAndDigit.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndDigit) &&
                    check(paragraphFormat.Alignment.CompareTo(defaultParagraphFormat.Alignment)) &&
                    paragraphFormat.AutoAdjustRightIndent.Equals(defaultParagraphFormat.AutoAdjustRightIndent) &&
                    check(paragraphFormat.BaseLineAlignment.CompareTo(defaultParagraphFormat.BaseLineAlignment)) &&
                    paragraphFormat.Borders.AlwaysInFront.Equals(defaultParagraphFormat.Borders.AlwaysInFront) &&
                    paragraphFormat.Borders.Count.Equals(defaultParagraphFormat.Borders.Count) &&
                    paragraphFormat.Borders.Enable.Equals(defaultParagraphFormat.Borders.Enable) &&
                    paragraphFormat.Borders.EnableFirstPageInSection.Equals(defaultParagraphFormat.Borders.EnableFirstPageInSection) &&
                    paragraphFormat.Borders.EnableOtherPagesInSection.Equals(defaultParagraphFormat.Borders.EnableOtherPagesInSection) &&

                    check(paragraphFormat.Borders.DistanceFrom.CompareTo(defaultParagraphFormat.Borders.DistanceFrom)) &&
                    check(paragraphFormat.Borders.DistanceFromBottom.CompareTo(defaultParagraphFormat.Borders.DistanceFromBottom)) &&
                    check(paragraphFormat.Borders.DistanceFromLeft.CompareTo(defaultParagraphFormat.Borders.DistanceFromLeft)) &&
                    check(paragraphFormat.Borders.DistanceFromRight.CompareTo(defaultParagraphFormat.Borders.DistanceFromRight)) &&
                    check(paragraphFormat.Borders.DistanceFromTop.CompareTo(defaultParagraphFormat.Borders.DistanceFromTop)) &&

                    paragraphFormat.Borders.HasHorizontal.Equals(defaultParagraphFormat.Borders.HasHorizontal) &&
                    paragraphFormat.Borders.HasVertical.Equals(defaultParagraphFormat.Borders.HasVertical) &&
                    check(paragraphFormat.Borders.InsideColor.CompareTo(defaultParagraphFormat.Borders.InsideColor)) &&
                    check(paragraphFormat.Borders.InsideColorIndex.CompareTo(defaultParagraphFormat.Borders.InsideColorIndex)) &&
                    check(paragraphFormat.Borders.InsideLineStyle.CompareTo(defaultParagraphFormat.Borders.InsideLineStyle)) &&
                    check(paragraphFormat.Borders.InsideLineWidth.CompareTo(defaultParagraphFormat.Borders.InsideLineWidth)) &&
                    paragraphFormat.Borders.JoinBorders.Equals(defaultParagraphFormat.Borders.JoinBorders) &&
                    check(paragraphFormat.Borders.OutsideColor.CompareTo(defaultParagraphFormat.Borders.OutsideColor)) &&
                    check(paragraphFormat.Borders.OutsideColorIndex.CompareTo(defaultParagraphFormat.Borders.OutsideColorIndex)) &&
                    check(paragraphFormat.Borders.OutsideLineStyle.CompareTo(defaultParagraphFormat.Borders.OutsideLineStyle)) &&
                    check(paragraphFormat.Borders.OutsideLineWidth.CompareTo(defaultParagraphFormat.Borders.OutsideLineWidth)) &&
                    paragraphFormat.Borders.Shadow.Equals(defaultParagraphFormat.Borders.Shadow) &&
                    paragraphFormat.Borders.SurroundFooter.Equals(defaultParagraphFormat.Borders.SurroundFooter) &&
                    paragraphFormat.Borders.SurroundHeader.Equals(defaultParagraphFormat.Borders.SurroundHeader) &&

                    //Возвращает или задает значение (в символах) для первого или висячего отступа. 
                    //Используйте положительное значение для установки отступа первой строки и используйте отрицательное значение, чтобы установить висячий отступ.
                    paragraphFormat.CharacterUnitFirstLineIndent.Equals(defaultParagraphFormat.CharacterUnitFirstLineIndent) &&

                    //Возвращает или задает левое значение отступа (в символах "ЗНАКАХ!!!!!") для указанных абзацев.(Также для правого "другая строчка снизу")
                    paragraphFormat.CharacterUnitLeftIndent.Equals(defaultParagraphFormat.CharacterUnitLeftIndent) &&
                    paragraphFormat.CharacterUnitRightIndent.Equals(defaultParagraphFormat.CharacterUnitRightIndent) &&

                   paragraphFormat.DisableLineHeightGrid.Equals(defaultParagraphFormat.DisableLineHeightGrid) &&

                   //Истинно, если Microsoft Word применяет восточноазиатские правила нарушения правил к указанным параграфам
                   paragraphFormat.FarEastLineBreakControl.Equals(defaultParagraphFormat.FarEastLineBreakControl) &&

                   //Отсуп красной строки, помоему в точках
                   paragraphFormat.FirstLineIndent.Equals(defaultParagraphFormat.FirstLineIndent) &&

                   //Истина, если Microsoft Word меняет знаки пунктуации в начале строки на символы полуширины для указанных абзацев.
                   paragraphFormat.HalfWidthPunctuationOnTopOfLine.Equals(defaultParagraphFormat.HalfWidthPunctuationOnTopOfLine) &&

                   //Истинно, если для указанных абзацев включена вешающая пунктуация. 
                   paragraphFormat.HangingPunctuation.Equals(defaultParagraphFormat.HangingPunctuation));
        }

        private bool check(int number)
        {
            return (number == 0); 
        }
    }
}
