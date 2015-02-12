/*
 * ParagraphExtensions.cs - Extensions for Paragraph
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
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    ///
    /// </summary>
    public static class ParagraphExtensions
    {
        public static readonly string DefaultParagraphStyleId = "Normal";
        public static readonly string DefaultCharacterStyleId = "DefaultParagraphFont";
        public static readonly int BodyTextOutlineLevel = 10;

        public static string GetParagraphStyleId(this Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
                return paragraph.ParagraphProperties.GetParagraphStyleId();

            return DefaultParagraphStyleId;
        }

        public static string GetParagraphStyleId(this ParagraphProperties pPr)
        {
            if (pPr.ParagraphStyleId != null && pPr.ParagraphStyleId.Val != null)
                return pPr.ParagraphStyleId.Val.Value;

            return DefaultParagraphStyleId;
        }

        public static string GetCharacterStyleId(this Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
                return paragraph.ParagraphProperties.GetCharacterStyleId();

            return DefaultCharacterStyleId;
        }

        public static string GetCharacterStyleId(this ParagraphProperties pPr)
        {
            if (pPr.ParagraphMarkRunProperties != null)
            {
                RunStyle rStyle = pPr.ParagraphMarkRunProperties.GetFirstChild<RunStyle>();
                if (rStyle != null && rStyle.Val != null)
                    return rStyle.Val.Value;
            }
            return DefaultCharacterStyleId;
        }

        public static Style GetParagraphStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(paragraph.GetParagraphStyleId());
        }

        public static Style GetParagraphStyle(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(pPr.GetParagraphStyleId());
        }

        public static Style GetCharacterStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            string styleId = paragraph.GetCharacterStyleId();
            if (styleId != null)
                return document.GetCharacterStyle(styleId);

            return null;
        }

        public static Style GetCharacterStyle(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            string styleId = pPr.GetCharacterStyleId();
            if (styleId != null)
                return document.GetCharacterStyle(styleId);

            return null;
        }

        public static int GetOutlineLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Style style = document.GetParagraphStyle(paragraph.GetParagraphStyleId());
            if (style != null && style.StyleParagraphProperties != null && style.StyleParagraphProperties.OutlineLevel != null)
                return style.StyleParagraphProperties.OutlineLevel.Val + 1;
            else
                return BodyTextOutlineLevel;
        }

        public static int GetListLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            ParagraphProperties pPr = paragraph.ParagraphProperties;
            if (pPr != null)
            {
                // Check whether paragraph has numbering properties.
                NumberingProperties numPr = pPr.NumberingProperties;
                if (numPr != null)
                {
                    NumberingLevelReference ilvl = numPr.NumberingLevelReference;
                    if (ilvl != null)
                        return ilvl.Val + 1;
                    else
                        return 1;
                }
                else
                {
                    // Check whether the style has numbering properties.
                    Style style = paragraph.GetParagraphStyle(document);
                    StyleParagraphProperties spPr = style.StyleParagraphProperties;
                    if (spPr != null)
                    {
                        numPr = spPr.NumberingProperties;
                        if (numPr != null)
                        {
                            NumberingLevelReference ilvl = numPr.NumberingLevelReference;
                            if (ilvl != null)
                                return ilvl.Val + 1;
                            else
                                return 1;
                        }
                    }
                }
            }

            // No numbering properties found.
            return 0;
        }

        public static T GetPropertiesLeafElement<T>(this Paragraph paragraph, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            // See whether the paragraph properties have this leaf element.
            ParagraphProperties pPr = paragraph.ParagraphProperties;
            if (pPr != null)
            {
                T leaf = pPr.Descendants<T>().FirstOrDefault();
                if (leaf != null)
                    return leaf;
            }

            // Don't look further if no document was given.
            if (document == null)
                return null;

            // See whether the paragraph style has it.
            Style style = paragraph.GetParagraphStyle(document);
            if (style != null)
            {
                T leaf = style.GetLeafElement<T>();
                if (leaf != null)
                    return leaf;
            }

            // See whether a potential character style has it.
            Style rStyle = paragraph.GetCharacterStyle(document);
            if (rStyle != null)
            {
                T leaf = rStyle.Descendants<T>().FirstOrDefault();
                if (leaf != null)
                    return leaf;
            }

            // There is no such leaf.
            return null;
        }

        public static bool IsBold(this Paragraph paragraph, WordprocessingDocument document)
        {
            Bold b = paragraph.GetPropertiesLeafElement<Bold>(document);
            if (b == null)
                return false;

            return b.Val == null ? true : b.Val.Value;
        }

        public static bool IsItalic(this Paragraph paragraph, WordprocessingDocument document)
        {
            Italic i = paragraph.GetPropertiesLeafElement<Italic>(document);
            if (i == null)
                return false;

            return i.Val == null ? true : i.Val.Value;
        }

        public static bool IsUnderline(this Paragraph paragraph, WordprocessingDocument document)
        {
            Underline u = paragraph.GetPropertiesLeafElement<Underline>(document);
            if (u == null)
                return false;

            return u.Val == null ? true : u.Val.Value == UnderlineValues.Single;
        }

        public static bool IsKeepNext(this Paragraph paragraph, WordprocessingDocument document)
        {
            KeepNext keepNext = paragraph.GetPropertiesLeafElement<KeepNext>(document);
            if (keepNext == null)
                return false;

            return keepNext.Val == null ? true : keepNext.Val.Value;
        }

        public static bool IsKeepNext(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            KeepNext keepNext = pPr.GetLeafElement<KeepNext>(document);
            if (keepNext == null)
                return false;

            return keepNext.Val == null ? true : keepNext.Val.Value;
        }

        public static T GetLeafElement<T>(this Style style)
            where T : OpenXmlLeafElement
        {
            T leaf = style.Descendants<T>().FirstOrDefault();
            if (leaf != null)
                return leaf;

            if (style.BasedOn != null)
            {
                Style baseStyle = style.Parent.Elements<Style>()
                    .FirstOrDefault(e => e.StyleId.Value == style.BasedOn.Val.Value);
                if (baseStyle != null)
                    return baseStyle.GetLeafElement<T>();
            }

            return null;
        }

        public static T GetLeafElement<T>(this ParagraphProperties pPr, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            // See whether the paragraph properties have this leaf element.
            T leaf = pPr.Descendants<T>().FirstOrDefault();
            if (leaf != null)
                return leaf;

            // Don't look further if no document was given.
            if (document == null)
                return null;

            // See whether the paragraph style has it.
            Style style = pPr.GetParagraphStyle(document);
            if (style != null)
            {
                leaf = style.GetLeafElement<T>();
                if (leaf != null)
                    return leaf;
            }

            // See whether a potential character style has it.
            Style rStyle = pPr.GetCharacterStyle(document);
            if (rStyle != null)
            {
                leaf = rStyle.Descendants<T>().FirstOrDefault();
                if (leaf != null)
                    return leaf;
            }

            // There is no such leaf.
            return null;
        }

        #region Trimming

        public static Paragraph TrimStart(this Paragraph paragraph, params char[] trimChars)
        {
            Paragraph p = (Paragraph)paragraph.CloneNode(true);

            bool continueTrimming = true;
            while (continueTrimming)
            {
                Text t = p.Descendants<Text>().FirstOrDefault();

                continueTrimming = t != null && t.Parent.Parent is Paragraph;
                if (continueTrimming)
                {
                    t.Text = t.Text.TrimStart(trimChars);

                    continueTrimming = t.Text == string.Empty;
                    if (continueTrimming)
                        t.Parent.Remove();
                }
            }

            return p;
        }

        public static Paragraph TrimEnd(this Paragraph paragraph, params char[] trimChars)
        {
            Paragraph p = (Paragraph)paragraph.CloneNode(true);

            bool continueTrimming = true;
            while (continueTrimming)
            {
                Text t = p.Descendants<Text>().LastOrDefault();

                continueTrimming = t != null && t.Parent.Parent is Paragraph;
                if (continueTrimming)
                {
                    t.Text = t.Text.TrimEnd(trimChars);

                    continueTrimming = t.Text == string.Empty;
                    if (continueTrimming)
                        t.Parent.Remove();
                }
            }

            return p;
        }

        public static Paragraph Trim(this Paragraph paragraph, params char[] trimChars)
        {
            return paragraph.TrimStart(trimChars).TrimEnd(trimChars);
        }

        public static Paragraph Trim(this Paragraph paragraph)
        {
            return paragraph.Trim(' ');
        }

        #endregion
    }
}
