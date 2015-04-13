/*
 * RunExtensions.cs - Extensions for Run
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
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public static class RunExtensions
    {
        public static Style GetCharacterStyle(this Run run, WordprocessingDocument document)
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.GetCharacterStyle(document);
        }

        public static string GetCharacterStyleId(this Run run)
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.GetCharacterStyleId();
        }

        public static T GetLeafElement<T>(this Run run, WordprocessingDocument document)
            where T : OpenXmlLeafElement
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.GetLeafElement<T>(document);
        }

        public static OnOffValue GetOnOffValue<T>(this Run run, WordprocessingDocument document)
            where T : OnOffType
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.GetOnOffValue<T>(document);
        }

        public static bool Is<T>(this Run run, WordprocessingDocument document)
            where T : OnOffType
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.Is<T>(document);
        }

        public static bool IsUnderlineSingle(this Run run, WordprocessingDocument document)
        {
            if (run == null)
                throw new ArgumentNullException("run");

            return run.RunProperties.IsUnderlineSingle(document);
        }
    }
}