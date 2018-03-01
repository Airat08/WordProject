using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
    class ParagraphFormat
    {
        private Word.ParagraphFormat defaultParagraphFormat;

        public ParagraphFormat(Word.ParagraphFormat defaultParagraphFormat)
        {
            this.defaultParagraphFormat = defaultParagraphFormat;
        }

        public List<int> checkingProperties(Word.ParagraphFormat paragraphFormat)
        {
            List<int> list = new List<int>();
            if (!paragraphFormat.AddSpaceBetweenFarEastAndAlpha.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndAlpha)) list.Add(1);
            if (!paragraphFormat.AddSpaceBetweenFarEastAndDigit.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndDigit)) list.Add(2);
            if (!check(paragraphFormat.Alignment.CompareTo(defaultParagraphFormat.Alignment))) list.Add(3);
            if (!paragraphFormat.AutoAdjustRightIndent.Equals(defaultParagraphFormat.AutoAdjustRightIndent)) list.Add(4);
            if (!check(paragraphFormat.BaseLineAlignment.CompareTo(defaultParagraphFormat.BaseLineAlignment))) list.Add(5);
            if (!paragraphFormat.Borders.AlwaysInFront.Equals(defaultParagraphFormat.Borders.AlwaysInFront)) list.Add(6);
            if (!paragraphFormat.Borders.Count.Equals(defaultParagraphFormat.Borders.Count)) list.Add(7);
            if (!paragraphFormat.Borders.Enable.Equals(defaultParagraphFormat.Borders.Enable)) list.Add(8);
            if (!paragraphFormat.Borders.EnableFirstPageInSection.Equals(defaultParagraphFormat.Borders.EnableFirstPageInSection)) list.Add(9);
            if (!paragraphFormat.Borders.EnableOtherPagesInSection.Equals(defaultParagraphFormat.Borders.EnableOtherPagesInSection)) list.Add(10);
            if (!check(paragraphFormat.Borders.DistanceFrom.CompareTo(defaultParagraphFormat.Borders.DistanceFrom))) list.Add(11);
            if (!check(paragraphFormat.Borders.DistanceFromBottom.CompareTo(defaultParagraphFormat.Borders.DistanceFromBottom))) list.Add(12);
            if (!check(paragraphFormat.Borders.DistanceFromLeft.CompareTo(defaultParagraphFormat.Borders.DistanceFromLeft))) list.Add(13);
            if (!check(paragraphFormat.Borders.DistanceFromRight.CompareTo(defaultParagraphFormat.Borders.DistanceFromRight))) list.Add(14);
            if (!check(paragraphFormat.Borders.DistanceFromTop.CompareTo(defaultParagraphFormat.Borders.DistanceFromTop))) list.Add(15);
            if (!paragraphFormat.Borders.HasHorizontal.Equals(defaultParagraphFormat.Borders.HasHorizontal)) list.Add(16);
            if (!paragraphFormat.Borders.HasVertical.Equals(defaultParagraphFormat.Borders.HasVertical)) list.Add(17);
            if (!check(paragraphFormat.Borders.InsideColor.CompareTo(defaultParagraphFormat.Borders.InsideColor))) list.Add(18);
            if (!check(paragraphFormat.Borders.InsideColorIndex.CompareTo(defaultParagraphFormat.Borders.InsideColorIndex))) list.Add(19);
            if (!check(paragraphFormat.Borders.InsideLineStyle.CompareTo(defaultParagraphFormat.Borders.InsideLineStyle))) list.Add(20);
            if (!check(paragraphFormat.Borders.InsideLineWidth.CompareTo(defaultParagraphFormat.Borders.InsideLineWidth))) list.Add(21);
            if (!paragraphFormat.Borders.JoinBorders.Equals(defaultParagraphFormat.Borders.JoinBorders)) list.Add(22);
            if (!check(paragraphFormat.Borders.OutsideColor.CompareTo(defaultParagraphFormat.Borders.OutsideColor))) list.Add(23);
            if (!check(paragraphFormat.Borders.OutsideColorIndex.CompareTo(defaultParagraphFormat.Borders.OutsideColorIndex))) list.Add(24);
            if (!check(paragraphFormat.Borders.OutsideLineStyle.CompareTo(defaultParagraphFormat.Borders.OutsideLineStyle))) list.Add(25);
            if (!check(paragraphFormat.Borders.OutsideLineWidth.CompareTo(defaultParagraphFormat.Borders.OutsideLineWidth))) list.Add(26);
            if (!paragraphFormat.Borders.Shadow.Equals(defaultParagraphFormat.Borders.Shadow)) list.Add(27);
            if (!paragraphFormat.Borders.SurroundFooter.Equals(defaultParagraphFormat.Borders.SurroundFooter)) list.Add(28);
            if (!paragraphFormat.Borders.SurroundHeader.Equals(defaultParagraphFormat.Borders.SurroundHeader)) list.Add(29);
            if (!paragraphFormat.CharacterUnitFirstLineIndent.Equals(defaultParagraphFormat.CharacterUnitFirstLineIndent)) list.Add(30);
            if (!paragraphFormat.CharacterUnitLeftIndent.Equals(defaultParagraphFormat.CharacterUnitLeftIndent)) list.Add(31);
            if (!paragraphFormat.CharacterUnitRightIndent.Equals(defaultParagraphFormat.CharacterUnitRightIndent)) list.Add(32);
            if (!paragraphFormat.DisableLineHeightGrid.Equals(defaultParagraphFormat.DisableLineHeightGrid)) list.Add(33);
            if (!paragraphFormat.FarEastLineBreakControl.Equals(defaultParagraphFormat.FarEastLineBreakControl)) list.Add(34);
            if (!paragraphFormat.FirstLineIndent.Equals(defaultParagraphFormat.FirstLineIndent)) list.Add(35);
            if (!paragraphFormat.HalfWidthPunctuationOnTopOfLine.Equals(defaultParagraphFormat.HalfWidthPunctuationOnTopOfLine)) list.Add(36);
            if (!paragraphFormat.HangingPunctuation.Equals(defaultParagraphFormat.HangingPunctuation)) list.Add(37);
            if (!paragraphFormat.Hyphenation.Equals(defaultParagraphFormat.Hyphenation)) list.Add(38);
            if (!paragraphFormat.KeepTogether.Equals(defaultParagraphFormat.KeepTogether)) list.Add(39);
            if (!paragraphFormat.LeftIndent.Equals(defaultParagraphFormat.LeftIndent)) list.Add(40);
            if (!paragraphFormat.LineSpacing.Equals(defaultParagraphFormat.LineSpacing)) list.Add(41);
            if (!paragraphFormat.LineSpacingRule.Equals(defaultParagraphFormat.LineSpacingRule)) list.Add(42);
            if (!paragraphFormat.LineUnitAfter.Equals(defaultParagraphFormat.LineUnitAfter)) list.Add(43);
            if (!paragraphFormat.LineUnitBefore.Equals(defaultParagraphFormat.LineUnitBefore)) list.Add(44);
            if (!paragraphFormat.MirrorIndents.Equals(defaultParagraphFormat.MirrorIndents)) list.Add(45);
            if (!paragraphFormat.NoLineNumber.Equals(defaultParagraphFormat.NoLineNumber)) list.Add(46);
            if (!paragraphFormat.OutlineLevel.Equals(defaultParagraphFormat.OutlineLevel)) list.Add(47);
            if (!paragraphFormat.PageBreakBefore.Equals(defaultParagraphFormat.PageBreakBefore)) list.Add(48);
            if (!paragraphFormat.RightIndent.Equals(defaultParagraphFormat.RightIndent)) list.Add(49);
            if (!check(paragraphFormat.Shading.BackgroundPatternColor.CompareTo(defaultParagraphFormat.Shading.BackgroundPatternColor))) list.Add(50);
            if (!check(paragraphFormat.Shading.BackgroundPatternColorIndex.CompareTo(defaultParagraphFormat.Shading.BackgroundPatternColorIndex))) list.Add(51);
            if (!check(paragraphFormat.Shading.ForegroundPatternColor.CompareTo(defaultParagraphFormat.Shading.ForegroundPatternColor))) list.Add(52);
            if (!check(paragraphFormat.Shading.ForegroundPatternColorIndex.CompareTo(defaultParagraphFormat.Shading.ForegroundPatternColorIndex))) list.Add(53);
            if (!check(paragraphFormat.Shading.Texture.CompareTo(defaultParagraphFormat.Shading.Texture))) list.Add(54);
            if (!paragraphFormat.SpaceAfter.Equals(defaultParagraphFormat.SpaceAfter)) list.Add(55);
            if (!paragraphFormat.SpaceAfterAuto.Equals(defaultParagraphFormat.SpaceAfterAuto)) list.Add(56);
            if (!paragraphFormat.SpaceBefore.Equals(defaultParagraphFormat.SpaceBefore)) list.Add(57);
            if (!paragraphFormat.SpaceBeforeAuto.Equals(defaultParagraphFormat.SpaceBeforeAuto)) list.Add(58);
            if (!paragraphFormat.TextboxTightWrap.Equals(defaultParagraphFormat.TextboxTightWrap)) list.Add(59);

            return list;
        }

        public bool CompareTo(Word.ParagraphFormat paragraphFormat)
        {
            return (paragraphFormat.AddSpaceBetweenFarEastAndAlpha.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndAlpha) &&
                    paragraphFormat.AddSpaceBetweenFarEastAndDigit.Equals(defaultParagraphFormat.AddSpaceBetweenFarEastAndDigit) &&
                    check(paragraphFormat.Alignment.CompareTo(defaultParagraphFormat.Alignment)) &&
                    paragraphFormat.AutoAdjustRightIndent.Equals(defaultParagraphFormat.AutoAdjustRightIndent) &&
                    check(paragraphFormat.BaseLineAlignment.CompareTo(defaultParagraphFormat.BaseLineAlignment)) &&

            #region Borders
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
            #endregion

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
                   paragraphFormat.HangingPunctuation.Equals(defaultParagraphFormat.HangingPunctuation) &&
                   paragraphFormat.Hyphenation.Equals(defaultParagraphFormat.Hyphenation) &&
                   paragraphFormat.KeepTogether.Equals(defaultParagraphFormat.KeepTogether) &&
                   paragraphFormat.LeftIndent.Equals(defaultParagraphFormat.LeftIndent) &&
                   paragraphFormat.LineSpacing.Equals(defaultParagraphFormat.LineSpacing) &&
                   paragraphFormat.LineSpacingRule.Equals(defaultParagraphFormat.LineSpacingRule) &&
                   paragraphFormat.LineUnitAfter.Equals(defaultParagraphFormat.LineUnitAfter) &&
                   paragraphFormat.LineUnitBefore.Equals(defaultParagraphFormat.LineUnitBefore) &&
                   paragraphFormat.MirrorIndents.Equals(defaultParagraphFormat.MirrorIndents) &&
                   paragraphFormat.NoLineNumber.Equals(defaultParagraphFormat.NoLineNumber) &&
                   paragraphFormat.OutlineLevel.Equals(defaultParagraphFormat.OutlineLevel) &&
                   paragraphFormat.PageBreakBefore.Equals(defaultParagraphFormat.PageBreakBefore) &&
                   paragraphFormat.RightIndent.Equals(defaultParagraphFormat.RightIndent) &&


            #region Shading
                   check(paragraphFormat.Shading.BackgroundPatternColor.CompareTo(defaultParagraphFormat.Shading.BackgroundPatternColor)) &&
                   check(paragraphFormat.Shading.BackgroundPatternColorIndex.CompareTo(defaultParagraphFormat.Shading.BackgroundPatternColorIndex)) &&
                   check(paragraphFormat.Shading.ForegroundPatternColor.CompareTo(defaultParagraphFormat.Shading.ForegroundPatternColor)) &&
                   check(paragraphFormat.Shading.ForegroundPatternColorIndex.CompareTo(defaultParagraphFormat.Shading.ForegroundPatternColorIndex)) &&
                   check(paragraphFormat.Shading.Texture.CompareTo(defaultParagraphFormat.Shading.Texture)) &&
            #endregion

                   paragraphFormat.SpaceAfter.Equals(defaultParagraphFormat.SpaceAfter) &&
                   paragraphFormat.SpaceAfterAuto.Equals(defaultParagraphFormat.SpaceAfterAuto) &&
                   paragraphFormat.SpaceBefore.Equals(defaultParagraphFormat.SpaceBefore) &&
                   paragraphFormat.SpaceBeforeAuto.Equals(defaultParagraphFormat.SpaceBeforeAuto) &&
                   paragraphFormat.TextboxTightWrap.Equals(defaultParagraphFormat.TextboxTightWrap)
                   );
        }

        private bool check(int number)
        {
            return (number == 0); 
        }
    }

}
