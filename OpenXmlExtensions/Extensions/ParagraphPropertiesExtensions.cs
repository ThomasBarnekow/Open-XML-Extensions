/*
 * ParagraphPropertiesExtensions.cs - Extensions for ParagraphProperties
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
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public static class ParagraphPropertiesExtensions
    {
        public static Style GetCharacterStyle(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var styleId = pPr.GetCharacterStyleId();
            return styleId != null ? document.GetCharacterStyle(styleId) : null;
        }

        public static string GetCharacterStyleId(this ParagraphProperties pPr)
        {
            if (pPr.ParagraphMarkRunProperties == null) return StyleExtensions.DefaultCharacterStyleId;

            var rStyle = pPr.ParagraphMarkRunProperties.GetFirstChild<RunStyle>();
            if (rStyle != null && rStyle.Val != null)
                return rStyle.Val.Value;

            return StyleExtensions.DefaultCharacterStyleId;
        }

        public static T GetLeafElement<T>(this ParagraphProperties pPr, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            // See whether the paragraph properties have this leaf element.
            var leaf = pPr.Descendants<T>().FirstOrDefault();
            if (leaf != null)
                return leaf;

            // Don't look further if no document was given.
            if (document == null)
                return null;

            // See whether the paragraph style has it.
            var style = pPr.GetParagraphStyle(document);
            if (style != null)
            {
                leaf = style.GetLeafElement<T>();
                if (leaf != null)
                    return leaf;
            }

            // See whether a potential character style has it.
            var rStyle = pPr.GetCharacterStyle(document);
            if (rStyle != null)
            {
                leaf = rStyle.Descendants<T>().FirstOrDefault();
                return leaf;
            }

            // There is no such leaf.
            return null;
        }

        public static Style GetParagraphStyle(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(pPr.GetParagraphStyleId());
        }

        public static string GetParagraphStyleId(this ParagraphProperties pPr)
        {
            if (pPr.ParagraphStyleId != null && pPr.ParagraphStyleId.Val != null)
                return pPr.ParagraphStyleId.Val.Value;

            return StyleExtensions.DefaultParagraphStyleId;
        }

        public static bool IsKeepNext(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            var keepNext = pPr.GetLeafElement<KeepNext>(document);
            return keepNext != null && (keepNext.Val == null || keepNext.Val.Value);
        }
    }
}