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

using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    public static class StyleExtensions
    {
        public const int BodyTextOutlineLevel = 10;
        public const string DefaultCharacterStyleId = "DefaultParagraphFont";
        public const string DefaultParagraphStyleId = "Normal";

        public static T GetLeafElement<T>(this Style style)
            where T : OpenXmlLeafElement
        {
            var leaf = style.Descendants<T>().FirstOrDefault();
            if (leaf != null) return leaf;

            if (style.BasedOn == null) return null;

            var baseStyle = style.Parent.Elements<Style>()
                .FirstOrDefault(e => e.StyleId.Value == style.BasedOn.Val.Value);

            return baseStyle != null ? baseStyle.GetLeafElement<T>() : null;
        }

        public static bool IsBold(this Style style)
        {
            var b = style.GetLeafElement<Bold>();
            return b != null && (b.Val == null || b.Val.Value);
        }

        public static bool IsItalic(this Style style)
        {
            var i = style.GetLeafElement<Italic>();
            return i != null && (i.Val == null || i.Val.Value);
        }

        public static bool IsUnderline(this Style style)
        {
            var u = style.GetLeafElement<Underline>();
            return u != null && u.Val != null && u.Val.Value == UnderlineValues.Single;
        }
    }
}