/*
 * RunPropertiesExtensions.cs - Extensions for RunProperties
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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
{
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public static class RunPropertiesExtensions
    {
        public static Style GetCharacterStyle(this RunProperties rPr, WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var styleId = rPr.GetCharacterStyleId();
            return styleId != null ? document.GetCharacterStyle(styleId) : null;
        }

        public static string GetCharacterStyleId(this RunProperties rPr)
        {
            return rPr != null && rPr.RunStyle != null ? rPr.RunStyle.Val.Value : null;
        }

        public static T GetLeafElement<T>(this RunProperties rPr, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            if (rPr == null) return null;

            var leafElement = rPr.GetFirstChild<T>();
            if (leafElement != null) return leafElement;

            var style = rPr.GetCharacterStyle(document);
            return style != null ? style.GetLeafElement<T>() : null;
        }

        public static OnOffValue GetOnOffValue<T>(this RunProperties rPr, WordprocessingDocument document)
            where T : OnOffType
        {
            var onOffElement = rPr.GetLeafElement<T>(document);
            if (onOffElement == null) return null;

            return onOffElement.Val ?? new OnOffValue(true);
        }

        public static bool Is<T>(this RunProperties rPr, WordprocessingDocument document)
            where T : OnOffType
        {
            var onOffValue = rPr.GetOnOffValue<T>(document);
            return onOffValue != null && onOffValue.Value;
        }

        public static bool IsUnderlineSingle(this RunProperties rPr, WordprocessingDocument document)
        {
            var u = rPr.GetLeafElement<Underline>(document);
            return u != null && u.Val != null && u.Val.Value == UnderlineValues.Single;
        }
    }
}