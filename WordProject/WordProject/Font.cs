using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordProject
{
    public class Font
    {
        private Word.Font defaultFont;
        private Word.Borders defaultBorders;
        private Word.FillFormat defaultFill;
        private Word.GlowFormat defaultGlow;
        private Word.LineFormat defaultLine;
        private Word.ReflectionFormat defaultReflection;
        private Word.Shading defaultShading;
        private Word.ShadowFormat defaultTextShadow;
        private Word.ThreeDFormat defaultThreeDFormat;

        public Font(Word.Font defaultFont)
        {
            this.defaultFont = defaultFont;
            this.defaultBorders = defaultFont.Borders;
            this.defaultFill = defaultFont.Fill;
            this.defaultGlow = defaultFont.Glow;
            this.defaultLine = defaultFont.Line;
            this.defaultReflection = defaultFont.Reflection;
            this.defaultShading = defaultFont.Shading;
            this.defaultTextShadow = defaultFont.TextShadow;
            this.defaultThreeDFormat = defaultFont.ThreeD;
        }

        public List<int> test_CompareTo(Word.Font font)
        {
            List<int> notEqList = new List<int>();

            if (!font.Color.Equals(Word.WdColor.wdColorAutomatic) && !font.Color.Equals(defaultFont.Color))
            {
                notEqList.Add(-1);
            }
            if (!font.ColorIndex.Equals(Word.WdColorIndex.wdAuto) && !font.ColorIndex.Equals(defaultFont.ColorIndex))
            {
                notEqList.Add(0);
            }
            if (!font.AllCaps.Equals(defaultFont.AllCaps)) { notEqList.Add(1); }
            if (!font.Bold.Equals(defaultFont.Bold)) { notEqList.Add(2); }
            #region Borders
            if (!font.Borders.AlwaysInFront.Equals(defaultBorders.AlwaysInFront)) { notEqList.Add(4); }
            if (!font.Borders.Count.Equals(defaultBorders.Count)) { notEqList.Add(5); }
            if (!font.Borders.DistanceFrom.Equals(defaultBorders.DistanceFrom)) { notEqList.Add(6); }
            //if (!font.Borders.DistanceFromBottom.Equals(defaultBorders.DistanceFromBottom)) { notEqList.Add(7); }
            //if (!font.Borders.DistanceFromLeft.Equals(defaultBorders.DistanceFromLeft)) { notEqList.Add(8); }
            //if (!font.Borders.DistanceFromRight.Equals(defaultBorders.DistanceFromRight)) { notEqList.Add(9); }
            //if (!font.Borders.DistanceFromTop.Equals(defaultBorders.DistanceFromTop)) { notEqList.Add(10); }
            if (!font.Borders.Enable.Equals(defaultBorders.Enable)) { notEqList.Add(11); }
            if (!font.Borders.EnableFirstPageInSection.Equals(defaultBorders.EnableFirstPageInSection)) { notEqList.Add(12); }
            if (!font.Borders.EnableOtherPagesInSection.Equals(defaultBorders.EnableOtherPagesInSection)) { notEqList.Add(13); }
            if (!font.Borders.HasHorizontal.Equals(defaultBorders.HasHorizontal)) { notEqList.Add(14); }
            if (!font.Borders.HasVertical.Equals(defaultBorders.HasVertical)) { notEqList.Add(15); }
            if (!font.Borders.InsideColor.Equals(defaultBorders.InsideColor)) { notEqList.Add(16); }
            if (!font.Borders.InsideColorIndex.Equals(defaultBorders.InsideColorIndex)) { notEqList.Add(17); }
            if (!font.Borders.InsideLineStyle.Equals(defaultBorders.InsideLineStyle)) { notEqList.Add(18); }
            if (!font.Borders.InsideLineWidth.Equals(defaultBorders.InsideLineWidth)) { notEqList.Add(19); }
            if (!font.Borders.JoinBorders.Equals(defaultBorders.JoinBorders)) { notEqList.Add(20); }
            if (!font.Borders.OutsideColor.Equals(defaultBorders.OutsideColor)) { notEqList.Add(21); }
            if (!font.Borders.OutsideColorIndex.Equals(defaultBorders.OutsideColorIndex)) { notEqList.Add(22); }
            if (!font.Borders.OutsideLineStyle.Equals(defaultBorders.OutsideLineStyle)) { notEqList.Add(23); }
            if (!font.Borders.OutsideLineWidth.Equals(defaultBorders.OutsideLineWidth)) { notEqList.Add(24); }
            if (!font.Borders.Shadow.Equals(defaultBorders.Shadow)) { notEqList.Add(25); }
            if (!font.Borders.SurroundFooter.Equals(defaultBorders.SurroundFooter)) { notEqList.Add(26); }
            if (!font.Borders.SurroundHeader.Equals(defaultBorders.SurroundHeader)) { notEqList.Add(27); }
            #endregion
            if (!font.ContextualAlternates.Equals(defaultFont.ContextualAlternates)) { notEqList.Add(29); }
            if (!font.DiacriticColor.Equals(defaultFont.DiacriticColor)) { notEqList.Add(30); }
            if (!font.DisableCharacterSpaceGrid.Equals(defaultFont.DisableCharacterSpaceGrid)) { notEqList.Add(31); }
            if (!font.DoubleStrikeThrough.Equals(defaultFont.DoubleStrikeThrough)) { notEqList.Add(32); }
            if (!font.Emboss.Equals(defaultFont.Emboss)) { notEqList.Add(33); }
            if (!font.EmphasisMark.Equals(defaultFont.EmphasisMark)) { notEqList.Add(34); }
            if (!font.Engrave.Equals(defaultFont.Engrave)) { notEqList.Add(35); }
            #region Fill
            ///!!!!!! !!!!! !!!! ColorFormat if (!font.Fill.BackColor.Equals(defaultFill.BackColor)) { notEqList.Add(37); }
            ////!!!! ! ! !  !!! ColorFormat if (!font.Fill.ForeColor.Equals(defaultFill.ForeColor)) { notEqList.Add(38); }
            //if (!font.Fill.GradientAngle.Equals(defaultFill.GradientAngle)) { notEqList.Add(39); }
            if (!font.Fill.GradientColorType.Equals(defaultFill.GradientColorType)) { notEqList.Add(40); }
            //if (!font.Fill.GradientDegree.Equals(defaultFill.GradientDegree)) { notEqList.Add(41); }
            //if (!font.Fill.GradientStops.Equals(defaultFill.GradientStops)) { notEqList.Add(42); }
            //if (!font.Fill.GradientStyle.Equals(defaultFill.GradientStyle)) { notEqList.Add(43); }
            //if (!font.Fill.GradientVariant.Equals(defaultFill.GradientVariant)) { notEqList.Add(44); }
            //if (!font.Fill.Pattern.Equals(defaultFill.Pattern)) { notEqList.Add(45); }
            //if (!font.Fill.PictureEffects.Equals(defaultFill.PictureEffects)) { notEqList.Add(46); }
            //if (!font.Fill.PresetGradientType.Equals(defaultFill.PresetGradientType)) { notEqList.Add(47); }
            //if (!font.Fill.PresetTexture.Equals(defaultFill.PresetTexture)) { notEqList.Add(48); }
            //if (!font.Fill.RotateWithObject.Equals(defaultFill.RotateWithObject)) { notEqList.Add(49); }
            //if (!font.Fill.TextureAlignment.Equals(defaultFill.TextureAlignment)) { notEqList.Add(50); }
            //if (!font.Fill.TextureHorizontalScale.Equals(defaultFill.TextureHorizontalScale)) { notEqList.Add(51); }
            //if (!font.Fill.TextureName.Equals(defaultFill.TextureName)) { notEqList.Add(52); }
            //if (!font.Fill.TextureOffsetX.Equals(defaultFill.TextureOffsetX)) { notEqList.Add(53); }
            //if (!font.Fill.TextureOffsetY.Equals(defaultFill.TextureOffsetY)) { notEqList.Add(54); }
            //if (!font.Fill.TextureTile.Equals(defaultFill.TextureTile)) { notEqList.Add(55); }
            //if (!font.Fill.TextureType.Equals(defaultFill.TextureType)) { notEqList.Add(56); }
            //if (!font.Fill.TextureVerticalScale.Equals(defaultFill.TextureVerticalScale)) { notEqList.Add(57); }
            if (!font.Fill.Transparency.Equals(defaultFill.Transparency)) { notEqList.Add(58); }
            if (!font.Fill.Type.Equals(defaultFill.Type)) { notEqList.Add(59); }
            if (!font.Fill.Visible.Equals(defaultFill.Visible)) { notEqList.Add(60); }
            #endregion
            #region Glow
            ////!!!!!!!! ColorFormat if (!font.Glow.Color.Equals(defaultGlow.Color)) { notEqList.Add(63); }
            if (!font.Glow.Radius.Equals(defaultGlow.Radius)) { notEqList.Add(64); }
            if (!font.Glow.Transparency.Equals(defaultGlow.Transparency)) { notEqList.Add(65); }
            #endregion
            if (!font.Hidden.Equals(defaultFont.Hidden)) { notEqList.Add(67); }
            if (!font.Italic.Equals(defaultFont.Italic)) { notEqList.Add(68); }
            if (!font.Kerning.Equals(defaultFont.Kerning)) { notEqList.Add(69); }
            if (!font.Ligatures.Equals(defaultFont.Ligatures)) { notEqList.Add(70); }
            #region Line
            //// !!!! ColorFormat if (!font.Line.BackColor.Equals(defaultLine.BackColor)) { notEqList.Add(72); }
            //if (!font.Line.BeginArrowheadLength.Equals(defaultLine.BeginArrowheadLength)) { notEqList.Add(73); }
            //if (!font.Line.BeginArrowheadStyle.Equals(defaultLine.BeginArrowheadStyle)) { notEqList.Add(74); }
            //if (!font.Line.BeginArrowheadWidth.Equals(defaultLine.BeginArrowheadWidth)) { notEqList.Add(75); }
            if (!font.Line.DashStyle.Equals(defaultLine.DashStyle)) { notEqList.Add(76); }
            //if (!font.Line.EndArrowheadLength.Equals(defaultLine.EndArrowheadLength)) { notEqList.Add(77); }
            //if (!font.Line.EndArrowheadStyle.Equals(defaultLine.EndArrowheadWidth)) { notEqList.Add(78); }
            ////!!!!!! ColorFormat if (!font.Line.ForeColor.Equals(defaultLine.ForeColor)) { notEqList.Add(79); }
            if (!font.Line.InsetPen.Equals(defaultLine.InsetPen)) { notEqList.Add(80); }
            //if (!font.Line.Pattern.Equals(defaultLine.Pattern)) { notEqList.Add(81); }
            if (!font.Line.Style.Equals(defaultLine.Style)) { notEqList.Add(82); }
            if (!font.Line.Transparency.Equals(defaultLine.Transparency)) { notEqList.Add(83); }
            if (!font.Line.Visible.Equals(defaultLine.Visible)) { notEqList.Add(84); }
            if (!font.Line.Weight.Equals(defaultLine.Weight)) { notEqList.Add(85); }
            #endregion
            if (!font.Name.Equals(defaultFont.Name)) { notEqList.Add(87); }
            if (!font.NameAscii.Equals(defaultFont.NameAscii)) { notEqList.Add(88); }
            if (!font.NameFarEast.Equals(defaultFont.NameFarEast)) { notEqList.Add(89); }
            if (!font.NameOther.Equals(defaultFont.NameOther)) { notEqList.Add(90); }
            if (!font.NumberForm.Equals(defaultFont.NumberForm)) { notEqList.Add(91); }
            if (!font.NumberSpacing.Equals(defaultFont.NumberSpacing)) { notEqList.Add(92); }
            if (!font.Outline.Equals(defaultFont.Outline)) { notEqList.Add(93); }
            #region Reflection
            if (!font.Reflection.Blur.Equals(defaultReflection.Blur)) { notEqList.Add(95); }
            if (!font.Reflection.Offset.Equals(defaultReflection.Offset)) { notEqList.Add(96); }
            if (!font.Reflection.Size.Equals(defaultReflection.Size)) { notEqList.Add(97); }
            if (!font.Reflection.Transparency.Equals(defaultReflection.Transparency)) { notEqList.Add(98); }
            if (!font.Reflection.Type.Equals(defaultReflection.Type)) { notEqList.Add(99); }
            #endregion
            if (!font.Scaling.Equals(defaultFont.Scaling)) { notEqList.Add(101); }
            #region Shaiding
            if (!font.Shading.BackgroundPatternColor.Equals(defaultShading.BackgroundPatternColor)) { notEqList.Add(103); }
            if (!font.Shading.BackgroundPatternColorIndex.Equals(defaultShading.BackgroundPatternColorIndex)) { notEqList.Add(104); }
            if (!font.Shading.ForegroundPatternColor.Equals(defaultShading.ForegroundPatternColor)) { notEqList.Add(105); }
            if (!font.Shading.ForegroundPatternColorIndex.Equals(defaultShading.ForegroundPatternColorIndex)) { notEqList.Add(106); }
            if (!font.Shading.Texture.Equals(defaultShading.Texture)) { notEqList.Add(107); }
            #endregion
            if (!font.Shadow.Equals(defaultFont.Shadow)) { notEqList.Add(109); }
            if (!font.Size.Equals(defaultFont.Size)) { notEqList.Add(110); }
            if (!font.SmallCaps.Equals(defaultFont.SmallCaps)) { notEqList.Add(111); }
            if (!font.Spacing.Equals(defaultFont.Spacing)) { notEqList.Add(112); }
            if (!font.StrikeThrough.Equals(defaultFont.StrikeThrough)) { notEqList.Add(113); }
            if (!font.StylisticSet.Equals(defaultFont.StylisticSet)) { notEqList.Add(114); }
            if (!font.Subscript.Equals(defaultFont.Subscript)) { notEqList.Add(115); }
            if (!font.Superscript.Equals(defaultFont.Superscript)) { notEqList.Add(116); }
            //// !!!!! ColorFormat if (!font.TextColor.Equals(defaultFont.TextColor)) { notEqList.Add(117); }
            #region TextShadow
            if (!font.TextShadow.Blur.Equals(defaultTextShadow.Blur)) { notEqList.Add(119); }
            /////!!!!! ColorFormat if (!font.TextShadow.ForeColor.Equals(defaultTextShadow.ForeColor)) { notEqList.Add(120); }
            if (!font.TextShadow.Obscured.Equals(defaultTextShadow.Obscured)) { notEqList.Add(121); }
            if (!font.TextShadow.OffsetX.Equals(defaultTextShadow.OffsetX)) { notEqList.Add(122); }
            if (!font.TextShadow.OffsetY.Equals(defaultTextShadow.OffsetY)) { notEqList.Add(123); }
            //if (!font.TextShadow.RotateWithShape.Equals(defaultTextShadow.RotateWithShape)) { notEqList.Add(124); }
            if (!font.TextShadow.Size.Equals(defaultTextShadow.Size)) { notEqList.Add(125); }
            if (!font.TextShadow.Style.Equals(defaultTextShadow.Style)) { notEqList.Add(126); }
            if (!font.TextShadow.Transparency.Equals(defaultTextShadow.Transparency)) { notEqList.Add(127); }
            if (!font.TextShadow.Type.Equals(defaultTextShadow.Type)) { notEqList.Add(128); }
            if (!font.TextShadow.Visible.Equals(defaultTextShadow.Visible)) { notEqList.Add(129); }
            #endregion
            #region ThreeD
            if (!font.ThreeD.BevelBottomDepth.Equals(defaultThreeDFormat.BevelBottomDepth)) { notEqList.Add(132); }
            if (!font.ThreeD.BevelBottomInset.Equals(defaultThreeDFormat.BevelBottomInset)) { notEqList.Add(133); }
            if (!font.ThreeD.BevelBottomType.Equals(defaultThreeDFormat.BevelBottomType)) { notEqList.Add(134); }
            if (!font.ThreeD.BevelTopDepth.Equals(defaultThreeDFormat.BevelTopDepth)) { notEqList.Add(135); }
            if (!font.ThreeD.BevelTopInset.Equals(defaultThreeDFormat.BevelTopInset)) { notEqList.Add(136); }
            if (!font.ThreeD.BevelTopType.Equals(defaultThreeDFormat.BevelTopType)) { notEqList.Add(137); }
            ////!!!!!! ColorFormat if (!font.ThreeD.ContourColor.Equals(defaultThreeDFormat.ContourColor)) { notEqList.Add(138); }
            if (!font.ThreeD.ContourWidth.Equals(defaultThreeDFormat.ContourWidth)) { notEqList.Add(139); }
            if (!font.ThreeD.Depth.Equals(defaultThreeDFormat.Depth)) { notEqList.Add(140); }
            ////!!!! Colorformat if (!font.ThreeD.ExtrusionColor.Equals(defaultThreeDFormat.ExtrusionColor)) { notEqList.Add(141); }
            if (!font.ThreeD.ExtrusionColorType.Equals(defaultThreeDFormat.ExtrusionColorType)) { notEqList.Add(142); }
            //if (!font.ThreeD.FieldOfView.Equals(defaultThreeDFormat.FieldOfView)) { notEqList.Add(143); }
            if (!font.ThreeD.LightAngle.Equals(defaultThreeDFormat.LightAngle)) { notEqList.Add(144); }
            if (!font.ThreeD.Perspective.Equals(defaultThreeDFormat.Perspective)) { notEqList.Add(145); }
            if (!font.ThreeD.PresetCamera.Equals(defaultThreeDFormat.PresetCamera)) { notEqList.Add(146); }
            if (!font.ThreeD.PresetExtrusionDirection.Equals(defaultThreeDFormat.PresetExtrusionDirection)) { notEqList.Add(147); }
            if (!font.ThreeD.PresetLighting.Equals(defaultThreeDFormat.PresetLighting)) { notEqList.Add(148); }
            if (!font.ThreeD.PresetLightingDirection.Equals(defaultThreeDFormat.PresetLightingDirection)) { notEqList.Add(149); }
            if (!font.ThreeD.PresetLightingSoftness.Equals(defaultThreeDFormat.PresetLightingSoftness)) { notEqList.Add(150); }
            if (!font.ThreeD.PresetMaterial.Equals(defaultThreeDFormat.PresetMaterial)) { notEqList.Add(151); }
            if (!font.ThreeD.PresetThreeDFormat.Equals(defaultThreeDFormat.PresetThreeDFormat)) { notEqList.Add(152); }
            //if (!font.ThreeD.ProjectText.Equals(defaultThreeDFormat.ProjectText)) { notEqList.Add(153); }
            //if (!font.ThreeD.RotationX.Equals(defaultThreeDFormat.RotationX)) { notEqList.Add(154); }
            //if (!font.ThreeD.RotationY.Equals(defaultThreeDFormat.RotationY)) { notEqList.Add(155); }
            //if (!font.ThreeD.RotationZ.Equals(defaultThreeDFormat.RotationZ)) { notEqList.Add(156); }
            if (!font.ThreeD.Visible.Equals(defaultThreeDFormat.Visible)) { notEqList.Add(157); }
            //if (!font.ThreeD.Z.Equals(defaultThreeDFormat.Z)) { notEqList.Add(158); }
            #endregion
            if (!font.Underline.Equals(defaultFont.Underline)) { notEqList.Add(160); }
            if (!font.UnderlineColor.Equals(defaultFont.UnderlineColor)){ notEqList.Add(161); }

            return notEqList;

        }

        public bool CompareTo(Word.Font font)
        {
            if (!font.Color.Equals(Word.WdColor.wdColorAutomatic) && !font.Color.Equals(defaultFont.Color))
            {
                return false;
            }
            if (!font.ColorIndex.Equals(Word.WdColorIndex.wdAuto) && !font.ColorIndex.Equals(defaultFont.ColorIndex))
            {
                return false;
            }
            return font.AllCaps.Equals(defaultFont.AllCaps)
                && font.Bold.Equals(defaultFont.Bold)
            #region Borders
                && font.Borders.AlwaysInFront.Equals(defaultBorders.AlwaysInFront)
                && font.Borders.Count.Equals(defaultBorders.Count)
                //&& font.Borders.DistanceFrom.Equals(defaultBorders.DistanceFrom)
                //&& font.Borders.DistanceFromBottom.Equals(defaultBorders.DistanceFromBottom)
                //&& font.Borders.DistanceFromLeft.Equals(defaultBorders.DistanceFromLeft)
                //&& font.Borders.DistanceFromRight.Equals(defaultBorders.DistanceFromRight)
                //&& font.Borders.DistanceFromTop.Equals(defaultBorders.DistanceFromTop)
                && font.Borders.Enable.Equals(defaultBorders.Enable)
                && font.Borders.EnableFirstPageInSection.Equals(defaultBorders.EnableFirstPageInSection)
                && font.Borders.EnableOtherPagesInSection.Equals(defaultBorders.EnableOtherPagesInSection)
                && font.Borders.HasHorizontal.Equals(defaultBorders.HasHorizontal)
                && font.Borders.HasVertical.Equals(defaultBorders.HasVertical)
                && font.Borders.InsideColor.Equals(defaultBorders.InsideColor)
                && font.Borders.InsideColorIndex.Equals(defaultBorders.InsideColorIndex)
                && font.Borders.InsideLineStyle.Equals(defaultBorders.InsideLineStyle)
                && font.Borders.InsideLineWidth.Equals(defaultBorders.InsideLineWidth)
                && font.Borders.JoinBorders.Equals(defaultBorders.JoinBorders)
                && font.Borders.OutsideColor.Equals(defaultBorders.OutsideColor)
                && font.Borders.OutsideColorIndex.Equals(defaultBorders.OutsideColorIndex)
                && font.Borders.OutsideLineStyle.Equals(defaultBorders.OutsideLineStyle)
                && font.Borders.OutsideLineWidth.Equals(defaultBorders.OutsideLineWidth)
                && font.Borders.Shadow.Equals(defaultBorders.Shadow)
                && font.Borders.SurroundFooter.Equals(defaultBorders.SurroundFooter)
                && font.Borders.SurroundHeader.Equals(defaultBorders.SurroundHeader)
            #endregion
                && font.ContextualAlternates.Equals(defaultFont.ContextualAlternates)
                && font.DiacriticColor.Equals(defaultFont.DiacriticColor)
                && font.DisableCharacterSpaceGrid.Equals(defaultFont.DisableCharacterSpaceGrid)
                && font.DoubleStrikeThrough.Equals(defaultFont.DoubleStrikeThrough)
                && font.Emboss.Equals(defaultFont.Emboss)
                && font.EmphasisMark.Equals(defaultFont.EmphasisMark)
                && font.Engrave.Equals(defaultFont.Engrave)
            #region Fill
                //&& font.Fill.BackColor.Equals(defaultFill.BackColor)
                //&& font.Fill.ForeColor.Equals(defaultFill.ForeColor)
                //&& font.Fill.GradientAngle.Equals(defaultFill.GradientAngle)
                && font.Fill.GradientColorType.Equals(defaultFill.GradientColorType)
                //&& font.Fill.GradientDegree.Equals(defaultFill.GradientDegree)
                //&& font.Fill.GradientStops.Equals(defaultFill.GradientStops)
                //&& font.Fill.GradientStyle.Equals(defaultFill.GradientStyle)
                //&& font.Fill.GradientVariant.Equals(defaultFill.GradientVariant)
                //&& font.Fill.Pattern.Equals(defaultFill.Pattern)
                //&& font.Fill.PictureEffects.Equals(defaultFill.PictureEffects)
                //&& font.Fill.PresetGradientType.Equals(defaultFill.PresetGradientType)
                //&& font.Fill.PresetTexture.Equals(defaultFill.PresetTexture)
                //&& font.Fill.RotateWithObject.Equals(defaultFill.RotateWithObject)
                //&& font.Fill.TextureAlignment.Equals(defaultFill.TextureAlignment)
                //&& font.Fill.TextureHorizontalScale.Equals(defaultFill.TextureHorizontalScale)
                //&& font.Fill.TextureName.Equals(defaultFill.TextureName)
                //&& font.Fill.TextureOffsetX.Equals(defaultFill.TextureOffsetX)
                //&& font.Fill.TextureOffsetY.Equals(defaultFill.TextureOffsetY)
                //&& font.Fill.TextureTile.Equals(defaultFill.TextureTile)
                //&& font.Fill.TextureType.Equals(defaultFill.TextureType)
                //&& font.Fill.TextureVerticalScale.Equals(defaultFill.TextureVerticalScale)
                && font.Fill.Transparency.Equals(defaultFill.Transparency)
                && font.Fill.Type.Equals(defaultFill.Type)
                && font.Fill.Visible.Equals(defaultFill.Visible)
            #endregion
            #region Glow
                && font.Glow.Color.Equals(defaultGlow.Color)
                && font.Glow.Radius.Equals(defaultGlow.Radius)
                && font.Glow.Transparency.Equals(defaultGlow.Transparency)
            #endregion
                && font.Hidden.Equals(defaultFont.Hidden)
                && font.Italic.Equals(defaultFont.Italic)
                && font.Kerning.Equals(defaultFont.Kerning)
                && font.Ligatures.Equals(defaultFont.Ligatures)
            #region Line
                && font.Line.BackColor.Equals(defaultLine.BackColor) ///!!
                && font.Line.BeginArrowheadLength.Equals(defaultLine.BeginArrowheadLength)
                && font.Line.BeginArrowheadStyle.Equals(defaultLine.BeginArrowheadStyle)
                && font.Line.BeginArrowheadWidth.Equals(defaultLine.BeginArrowheadWidth)
                && font.Line.DashStyle.Equals(defaultLine.DashStyle)
                && font.Line.EndArrowheadLength.Equals(defaultLine.EndArrowheadLength)
                && font.Line.EndArrowheadStyle.Equals(defaultLine.EndArrowheadWidth)
                && font.Line.ForeColor.Equals(defaultLine.ForeColor) ///!!!
                && font.Line.InsetPen.Equals(defaultLine.InsetPen)
                && font.Line.Pattern.Equals(defaultLine.Pattern)
                && font.Line.Style.Equals(defaultLine.Style)
                && font.Line.Transparency.Equals(defaultLine.Transparency)
                && font.Line.Visible.Equals(defaultLine.Visible)
                && font.Line.Weight.Equals(defaultLine.Weight)
            #endregion
                && font.Name.Equals(defaultFont.Name)
                && font.NameAscii.Equals(defaultFont.NameAscii)
                && font.NameFarEast.Equals(defaultFont.NameFarEast)
                && font.NameOther.Equals(defaultFont.NameOther)
                && font.NumberForm.Equals(defaultFont.NumberForm)
                && font.NumberSpacing.Equals(defaultFont.NumberSpacing)
                && font.Outline.Equals(defaultFont.Outline)
            #region Reflection
                && font.Reflection.Blur.Equals(defaultReflection.Blur)
                && font.Reflection.Offset.Equals(defaultReflection.Offset)
                && font.Reflection.Size.Equals(defaultReflection.Size)
                && font.Reflection.Transparency.Equals(defaultReflection.Transparency)
                && font.Reflection.Type.Equals(defaultReflection.Type)
            #endregion
                && font.Scaling.Equals(defaultFont.Scaling)
            #region Shaiding
                && font.Shading.BackgroundPatternColor.Equals(defaultShading.BackgroundPatternColor)
                && font.Shading.BackgroundPatternColorIndex.Equals(defaultShading.BackgroundPatternColorIndex)
                && font.Shading.ForegroundPatternColor.Equals(defaultShading.ForegroundPatternColor)
                && font.Shading.ForegroundPatternColorIndex.Equals(defaultShading.ForegroundPatternColorIndex)
                && font.Shading.Texture.Equals(defaultShading.Texture)
            #endregion
                && font.Shadow.Equals(defaultFont.Shadow)
                && font.Size.Equals(defaultFont.Size)
                && font.SmallCaps.Equals(defaultFont.SmallCaps)
                && font.Spacing.Equals(defaultFont.Spacing)
                && font.StrikeThrough.Equals(defaultFont.StrikeThrough)
                && font.StylisticSet.Equals(defaultFont.StylisticSet)
                && font.Subscript.Equals(defaultFont.Subscript)
                && font.Superscript.Equals(defaultFont.Superscript)
                //&& font.TextColor.Equals(defaultFont.TextColor)
            #region TextShadow
                && font.TextShadow.Blur.Equals(defaultTextShadow.Blur)
                && font.TextShadow.ForeColor.Equals(defaultTextShadow.ForeColor)
                && font.TextShadow.Obscured.Equals(defaultTextShadow.Obscured)
                && font.TextShadow.OffsetX.Equals(defaultTextShadow.OffsetX)
                && font.TextShadow.OffsetY.Equals(defaultTextShadow.OffsetY)
                && font.TextShadow.RotateWithShape.Equals(defaultTextShadow.RotateWithShape)
                && font.TextShadow.Size.Equals(defaultTextShadow.Size)
                && font.TextShadow.Style.Equals(defaultTextShadow.Style)
                && font.TextShadow.Transparency.Equals(defaultTextShadow.Transparency)
                && font.TextShadow.Type.Equals(defaultTextShadow.Type)
                && font.TextShadow.Visible.Equals(defaultTextShadow.Visible)
            #endregion
            #region ThreeD
                && font.ThreeD.BevelBottomDepth.Equals(defaultThreeDFormat.BevelBottomDepth)
                && font.ThreeD.BevelBottomInset.Equals(defaultThreeDFormat.BevelBottomInset)
                && font.ThreeD.BevelBottomType.Equals(defaultThreeDFormat.BevelBottomType)
                && font.ThreeD.BevelTopDepth.Equals(defaultThreeDFormat.BevelTopDepth)
                && font.ThreeD.BevelTopInset.Equals(defaultThreeDFormat.BevelTopInset)
                && font.ThreeD.BevelTopType.Equals(defaultThreeDFormat.BevelTopType)
                && font.ThreeD.ContourColor.Equals(defaultThreeDFormat.ContourColor)
                && font.ThreeD.ContourWidth.Equals(defaultThreeDFormat.ContourWidth)
                && font.ThreeD.Depth.Equals(defaultThreeDFormat.Depth)
                && font.ThreeD.ExtrusionColor.Equals(defaultThreeDFormat.ExtrusionColor)
                && font.ThreeD.ExtrusionColorType.Equals(defaultThreeDFormat.ExtrusionColorType)
                && font.ThreeD.FieldOfView.Equals(defaultThreeDFormat.FieldOfView)
                && font.ThreeD.LightAngle.Equals(defaultThreeDFormat.LightAngle)
                && font.ThreeD.Perspective.Equals(defaultThreeDFormat.Perspective)
                && font.ThreeD.PresetCamera.Equals(defaultThreeDFormat.PresetCamera)
                && font.ThreeD.PresetExtrusionDirection.Equals(defaultThreeDFormat.PresetExtrusionDirection)
                && font.ThreeD.PresetLighting.Equals(defaultThreeDFormat.PresetLighting)
                && font.ThreeD.PresetLightingDirection.Equals(defaultThreeDFormat.PresetLightingDirection)
                && font.ThreeD.PresetLightingSoftness.Equals(defaultThreeDFormat.PresetLightingSoftness)
                && font.ThreeD.PresetMaterial.Equals(defaultThreeDFormat.PresetMaterial)
                && font.ThreeD.PresetThreeDFormat.Equals(defaultThreeDFormat.PresetThreeDFormat)
                && font.ThreeD.ProjectText.Equals(defaultThreeDFormat.ProjectText)
                && font.ThreeD.RotationX.Equals(defaultThreeDFormat.RotationX)
                && font.ThreeD.RotationY.Equals(defaultThreeDFormat.RotationY)
                && font.ThreeD.RotationZ.Equals(defaultThreeDFormat.RotationZ)
                && font.ThreeD.Visible.Equals(defaultThreeDFormat.Visible)
                && font.ThreeD.Z.Equals(defaultThreeDFormat.Z)
            #endregion
                && font.Underline.Equals(defaultFont.Underline)
                && font.UnderlineColor.Equals(defaultFont.UnderlineColor);
        }
    }
}
