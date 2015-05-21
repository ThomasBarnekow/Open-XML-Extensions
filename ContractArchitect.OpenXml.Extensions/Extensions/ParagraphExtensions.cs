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

namespace ContractArchitect.OpenXml.Extensions
{
    /// <summary>
    /// Provides various extensions of the <see cref="Paragraph" /> class.
    /// </summary>
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public static class ParagraphExtensions
    {
        public static int GetIndentationLeft(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.ParagraphProperties.GetIndentationLeft(document);
        }

        public static int GetListLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            var numPr = paragraph.GetNumberingProperties(document);
            if (numPr == null)
                return 0;

            var numId = numPr.NumberingId;
            if (numId != null && numId.Val.Value == 0)
                return 0;

            var ilvl = numPr.NumberingLevelReference;
            if (ilvl != null)
                return ilvl.Val + 1;

            return 1;
        }

        public static int GetNumberedOutlineLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            // A paragraph without any numbering properties applied to it manually or by
            // assigning a style shall be on the body text outline level.
            var numPr = paragraph.GetNumberingProperties(document);
            if (numPr == null)
                return StyleExtensions.BodyTextOutlineLevel;

            // A paragraph the numbering of which has been turned off manually shall be
            // on the body text outline level as well.
            var numId = numPr.NumberingId;
            if (numId != null && numId.Val.Value == 0)
                return StyleExtensions.BodyTextOutlineLevel;

            // A numbered paragraph shall be on the outline level assigned to it by its style.
            return paragraph.GetOutlineLevel(document);
        }

        public static NumberingProperties GetNumberingProperties(this Paragraph paragraph,
            WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.ParagraphProperties.GetNumberingProperties(document);
        }

        public static string GetNumberingText(this Paragraph paragraph, NumberingState numberingState,
            WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.ParagraphProperties.GetNumberingText(numberingState, document);
        }

        public static int GetOutlineLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.GetParagraphStyle(document).GetOutlineLevel();
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

        public static Type GetRunTrackChangeType(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.RunsAreTrackChangeType<RunTrackChangeType>()
                ? paragraph.Descendants<Run>().First().Parent.GetType()
                : null;
        }

        public static bool HasStyleSeparator(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            return paragraph.ParagraphProperties.HasStyleSeparator();
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

        public static bool RunsAreTrackChangeType<T>(this Paragraph paragraph) where T : RunTrackChangeType
        {
            return paragraph != null &&
                   paragraph.Descendants<Run>().Any() &&
                   paragraph.Descendants<Run>().All(run => run.Parent is T);
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