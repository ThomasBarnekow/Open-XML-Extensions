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
        public static int GetListLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
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
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
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
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(paragraph.GetParagraphStyleId());
        }

        public static string GetParagraphStyleId(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.ParagraphProperties.GetParagraphStyleId();
        }

        public static bool Is<T>(this Paragraph paragraph, WordprocessingDocument document)
            where T : OnOffType
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            var style = paragraph.GetParagraphStyle(document);
            if (style.Is<Bold>())
            {
                var isAnyValueTurnedOff = paragraph.Descendants<Run>()
                    .Select(r => r.GetOnOffValue<T>(document))
                    .Any(val => val != null && val.Value == false);

                return !isAnyValueTurnedOff;
            }
            return paragraph.Descendants<Run>().All(r => r.Is<T>(document));
        }

        public static bool IsKeepNext(this Paragraph paragraph, WordprocessingDocument document)
        {
            return paragraph.ParagraphProperties.IsKeepNext(document);
        }

        public static bool IsUnderlineSingle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            var style = paragraph.GetParagraphStyle(document);
            if (style.IsUnderlineSingle())
            {
                var isUnderlineTurnedOff = paragraph.Descendants<Run>()
                    .Select(r => r.GetLeafElement<Underline>(document))
                    .Any(u => u.Val != null && u.Val.Value != UnderlineValues.Single);

                return !isUnderlineTurnedOff;
            }
            return paragraph.Descendants<Run>().All(r => r.IsUnderlineSingle(document));
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