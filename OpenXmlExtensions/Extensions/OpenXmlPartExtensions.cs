/*
 * OpenXmlPartExtensions.cs - Extensions for OpenXmlPart
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
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for the <see cref="OpenXmlPart" /> class.
    /// </summary>
    public static class OpenXmlPartExtensions
    {
        /// <summary>
        /// Returns the <see cref="OpenXmlPart" />'s root <see cref="XElement" />.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns>The root element.</returns>
        public static XElement GetRootElement(this OpenXmlPart part)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            try
            {
                using (var stream = part.GetStream())
                using (var streamReader = new StreamReader(stream))
                using (var xmlReader = XmlReader.Create(streamReader))
                    return XElement.Load(xmlReader);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the <see cref="OpenXmlPart" />'s root element <see cref="XName" />.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns>The name.</returns>
        public static XName GetRootName(this OpenXmlPart part)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            var root = part.GetRootElement();
            return root != null ? root.Name : null;
        }

        /// <summary>
        /// Returns the <see cref="OpenXmlPart" />'s root element <see cref="XNamespace" />.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns>The namespace.</returns>
        public static XNamespace GetRootNamespace(this OpenXmlPart part)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            var root = part.GetRootElement();
            return root != null ? root.Name.Namespace : null;
        }

        /// <summary>
        /// Sets the <see cref="OpenXmlPart" />'s root <see cref="XElement" />, replacing an
        /// existing one if it exists.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <param name="root">The new root element.</param>
        public static void SetRootElement(this OpenXmlPart part, XElement root)
        {
            if (part == null)
                throw new ArgumentNullException("part");
            if (root == null)
                throw new ArgumentNullException("root");

            using (var stream = part.GetStream(FileMode.Create))
            using (var streamWriter = new StreamWriter(stream))
            using (var xmlWriter = XmlWriter.Create(streamWriter))
                root.WriteTo(xmlWriter);
        }
    }
}
