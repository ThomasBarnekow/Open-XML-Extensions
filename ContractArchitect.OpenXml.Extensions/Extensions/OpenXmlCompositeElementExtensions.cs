/*
 * OpenXmlCompositeElementExtensions.cs - Extensions for OpenXmlCompositeElement
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
    public static class OpenXmlCompositeElementExtensions
    {
        public static T GetLeafElement<T>(this OpenXmlCompositeElement element)
            where T : OpenXmlLeafElement
        {
            return element == null ? null : element.GetFirstChild<T>();
        }

        public static OnOffValue GetOnOffValue<T>(this OpenXmlCompositeElement element)
            where T : OnOffType
        {
            var onOffElement = element.GetLeafElement<T>();
            if (onOffElement == null) return null;

            return onOffElement.Val ?? new OnOffValue(true);
        }

        public static bool Is<T>(this OpenXmlCompositeElement element)
            where T : OnOffType
        {
            var onOffValue = element.GetOnOffValue<T>();
            return onOffValue != null && onOffValue.Value;
        }
    }
}
