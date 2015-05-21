/*
 * StyleExtensions.cs - Extensions for Style
 *
 * Copyright 2014-2015 Thomas Barnekow
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * Developer: Thomas Barnekow
 * Email: thomas<at/>barnekow<dot/>info
 *
 * Version: 1.0.01
 */

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
{
    public static class StyleExtensions
    {
        public const int BodyTextOutlineLevel = 10;
        public const string DefaultCharacterStyleId = "DefaultParagraphFont";
        public const string DefaultParagraphStyleId = "Normal";

        public static Style GetBaseStyle(this Style style)
        {
            if (style == null)
                throw new ArgumentNullException("style");

            return style.BasedOn != null
                ? style.Parent.Elements<Style>().FirstOrDefault(e => e.StyleId.Value == style.BasedOn.Val.Value)
                : null;
        }

        public static int GetIndentationLeft(this Style style, WordprocessingDocument document)
        {
            if (style == null) return 0;

            var pPr = style.StyleParagraphProperties;
            if (pPr == null)
                return style.GetBaseStyle().GetIndentationLeft(document);

            var ind = pPr.Indentation;
            if (ind != null && ind.Left != null)
                return int.Parse(ind.Left.Value);

            var numPr = pPr.NumberingProperties;
            return numPr != null
                ? numPr.GetIndentationLeft(document)
                : style.GetBaseStyle().GetIndentationLeft(document);
        }

        public static NumberingProperties GetNumberingProperties(this Style style)
        {
            return style != null && style.StyleParagraphProperties != null
                ? style.StyleParagraphProperties.NumberingProperties
                : null;
        }

        public static string GetNumberingText(this Style style, NumberingState numberingState,
            WordprocessingDocument document)
        {
            if (style == null) return string.Empty;

            var pPr = style.StyleParagraphProperties;
            if (pPr == null)
                return style.GetBaseStyle().GetNumberingText(numberingState, document);

            var numPr = pPr.NumberingProperties;
            return numPr != null
                ? numPr.GetNumberingText(numberingState, document)
                : style.GetBaseStyle().GetNumberingText(numberingState, document);
        }

        public static T GetLeafElement<T>(this Style style)
            where T : OpenXmlLeafElement
        {
            if (style == null) return null;

            var leaf = style.Descendants<T>().FirstOrDefault();
            if (leaf != null) return leaf;

            var baseStyle = style.GetBaseStyle();

            return baseStyle != null ? baseStyle.GetLeafElement<T>() : null;
        }

        public static OnOffValue GetOnOffValue<T>(this Style style)
            where T : OnOffType
        {
            var onOffElement = style.GetLeafElement<T>();
            if (onOffElement == null) return null;

            return onOffElement.Val ?? new OnOffValue(true);
        }

        public static int GetOutlineLevel(this Style style)
        {
            if (style == null) return BodyTextOutlineLevel;

            if (style.StyleParagraphProperties != null && style.StyleParagraphProperties.OutlineLevel != null)
                return style.StyleParagraphProperties.OutlineLevel.Val + 1;

            var baseStyle = style.GetBaseStyle();

            return baseStyle != null ? baseStyle.GetOutlineLevel() : BodyTextOutlineLevel;
        }

        public static bool Is<T>(this Style style)
            where T : OnOffType
        {
            var onOffValue = style.GetOnOffValue<T>();
            return onOffValue != null && onOffValue.Value;
        }

        public static bool IsUnderlineSingle(this Style style)
        {
            var u = style.GetLeafElement<Underline>();
            return u != null && u.Val != null && u.Val.Value == UnderlineValues.Single;
        }
    }
}