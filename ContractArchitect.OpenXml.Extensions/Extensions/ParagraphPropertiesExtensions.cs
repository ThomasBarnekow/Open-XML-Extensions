﻿/*
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
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
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

        public static int GetIndentationLeft(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (pPr == null) return 0;

            var ind = pPr.Indentation;
            if (ind != null && ind.Left != null)
                return int.Parse(ind.Left.Value);

            var numPr = pPr.NumberingProperties;
            return numPr != null
                ? numPr.GetIndentationLeft(document)
                : pPr.GetParagraphStyle(document).GetIndentationLeft(document);
        }

        public static NumberingProperties GetNumberingProperties(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (pPr == null) return null;

            return pPr.NumberingProperties ?? pPr.GetParagraphStyle(document).GetNumberingProperties();
        }

        public static string GetNumberingText(this ParagraphProperties pPr, NumberingState numberingState,
            WordprocessingDocument document)
        {
            if (pPr == null) return string.Empty;

            var numPr = pPr.NumberingProperties;
            return numPr != null
                ? numPr.GetNumberingText(numberingState, document)
                : pPr.GetParagraphStyle(document).GetNumberingText(numberingState, document);
        }

        public static T GetLeafElement<T>(this ParagraphProperties pPr, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var leaf = pPr.Descendants<T>().FirstOrDefault();
            if (leaf != null) return leaf;

            var style = pPr.GetParagraphStyle(document);
            return style != null ? style.GetLeafElement<T>() : null;
        }

        public static OnOffValue GetOnOffValue<T>(this ParagraphProperties pPr, WordprocessingDocument document)
            where T : OnOffType
        {
            if (pPr == null)
                throw new ArgumentNullException("pPr");

            var element = pPr.Descendants<T>().FirstOrDefault();
            if (element != null)
                return element.Val.HasValue ? element.Val : new OnOffValue(true);

            var style = pPr.GetParagraphStyle(document);
            return style != null ? style.GetOnOffValue<T>() : null;
        }

        public static Style GetParagraphStyle(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(pPr.GetParagraphStyleId());
        }

        public static string GetParagraphStyleId(this ParagraphProperties pPr)
        {
            if (pPr == null || pPr.ParagraphStyleId == null)
                return StyleExtensions.DefaultParagraphStyleId;

            return pPr.ParagraphStyleId.Val.Value;
        }

        public static bool HasStyleSeparator(this ParagraphProperties pPr)
        {
            return pPr != null && pPr.Descendants<SpecVanish>().Any();
        }

        public static bool IsKeepNext(this ParagraphProperties pPr, WordprocessingDocument document)
        {
            if (pPr == null) return false;

            var val = pPr.GetOnOffValue<KeepNext>(document);
            return val != null && val.Value;
        }
    }
}