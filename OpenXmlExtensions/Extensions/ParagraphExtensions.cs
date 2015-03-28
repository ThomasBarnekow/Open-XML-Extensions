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
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    /// </summary>
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public static class ParagraphExtensions
    {
        public static Style GetCharacterStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var styleId = paragraph.GetCharacterStyleId();
            return styleId != null ? document.GetCharacterStyle(styleId) : null;
        }

        public static string GetCharacterStyleId(this Paragraph paragraph)
        {
            return paragraph.ParagraphProperties != null
                ? paragraph.ParagraphProperties.GetCharacterStyleId()
                : StyleExtensions.DefaultCharacterStyleId;
        }

        public static int GetListLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var pPr = paragraph.ParagraphProperties;
            if (pPr != null)
            {
                // Check whether paragraph has numbering properties.
                var numPr = pPr.NumberingProperties;
                if (numPr != null)
                {
                    var ilvl = numPr.NumberingLevelReference;
                    if (ilvl != null)
                        return ilvl.Val + 1;
                    return 1;
                }
                // Check whether the style has numbering properties.
                var style = paragraph.GetParagraphStyle(document);
                var spPr = style.StyleParagraphProperties;
                if (spPr != null)
                {
                    numPr = spPr.NumberingProperties;
                    if (numPr != null)
                    {
                        var ilvl = numPr.NumberingLevelReference;
                        if (ilvl != null)
                            return ilvl.Val + 1;
                        return 1;
                    }
                }
            }

            // No numbering properties found.
            return 0;
        }

        public static int GetOutlineLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var style = document.GetParagraphStyle(paragraph.GetParagraphStyleId());
            if (style != null && style.StyleParagraphProperties != null &&
                style.StyleParagraphProperties.OutlineLevel != null)
                return style.StyleParagraphProperties.OutlineLevel.Val + 1;

            return StyleExtensions.BodyTextOutlineLevel;
        }

        public static Style GetParagraphStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(paragraph.GetParagraphStyleId());
        }

        public static string GetParagraphStyleId(this Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
                return paragraph.ParagraphProperties.GetParagraphStyleId();

            return StyleExtensions.DefaultParagraphStyleId;
        }

        public static T GetPropertiesLeafElement<T>(this Paragraph paragraph, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            // See whether the paragraph properties have this leaf element.
            var pPr = paragraph.ParagraphProperties;
            if (pPr != null)
            {
                var leaf = pPr.Descendants<T>().FirstOrDefault();
                if (leaf != null)
                    return leaf;
            }

            // Don't look further if no document was given.
            if (document == null)
                return null;

            // See whether the paragraph style has it.
            var style = paragraph.GetParagraphStyle(document);
            if (style != null)
            {
                var leaf = style.GetLeafElement<T>();
                if (leaf != null)
                    return leaf;
            }

            // See whether a potential character style has it.
            var rStyle = paragraph.GetCharacterStyle(document);
            if (rStyle != null)
            {
                var leaf = rStyle.Descendants<T>().FirstOrDefault();
                return leaf;
            }

            // There is no such leaf.
            return null;
        }

        public static bool IsBold(this Paragraph paragraph, WordprocessingDocument document)
        {
            var b = paragraph.GetPropertiesLeafElement<Bold>(document);
            return b != null && (b.Val == null || b.Val.Value);
        }

        public static bool IsItalic(this Paragraph paragraph, WordprocessingDocument document)
        {
            var i = paragraph.GetPropertiesLeafElement<Italic>(document);
            return i != null && (i.Val == null || i.Val.Value);
        }

        public static bool IsKeepNext(this Paragraph paragraph, WordprocessingDocument document)
        {
            var keepNext = paragraph.GetPropertiesLeafElement<KeepNext>(document);
            return keepNext != null && (keepNext.Val == null || keepNext.Val.Value);
        }

        public static bool IsUnderline(this Paragraph paragraph, WordprocessingDocument document)
        {
            var u = paragraph.GetPropertiesLeafElement<Underline>(document);
            return u != null && u.Val != null && u.Val.Value == UnderlineValues.Single;
        }

        #region Trimming

        public static Paragraph TrimStart(this Paragraph paragraph, params char[] trimChars)
        {
            var p = (Paragraph) paragraph.CloneNode(true);

            var continueTrimming = true;
            while (continueTrimming)
            {
                var t = p.Descendants<Text>().FirstOrDefault();

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
            var p = (Paragraph) paragraph.CloneNode(true);

            var continueTrimming = true;
            while (continueTrimming)
            {
                var t = p.Descendants<Text>().LastOrDefault();

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