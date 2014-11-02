﻿/*
 * OpenXmlPartExtensions.cs - Extensions for OpenXmlPart
 * 
 * Copyright 2014 Thomas Barnekow
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
    /// Provides extension methods for the <see cref="OpenXmlPart"/> class.
    /// </summary>
    public static class OpenXmlPartExtensions
    {
        /// <summary>
        /// Returns the <see cref="OpenXmlPart"/>'s root <see cref="XElement"/>.
        /// </summary>
        /// <param name="part">The part</param>
        /// <returns>The root element</returns>
        public static XElement GetRootElement(this OpenXmlPart part)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            using (Stream stream = part.GetStream())
            using (StreamReader streamReader = new StreamReader(stream))
            using (XmlReader xmlReader = XmlReader.Create(streamReader))
                return XElement.Load(xmlReader);
        }

        /// <summary>
        /// Returns the <see cref="OpenXmlPart"/>'s root <see cref="XNamespace"/>.
        /// </summary>
        /// <param name="part">The part</param>
        /// <returns>The namespace</returns>
        public static XNamespace GetRootNamespace(this OpenXmlPart part)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            XElement root = GetRootElement(part);
            if (root != null)
                return root.Name.Namespace;
            else
                return null;
        }

        /// <summary>
        /// Sets the <see cref="OpenXmlPart"/>'s root <see cref="XElement"/>, replacing an
        /// existing one as necessary.
        /// </summary>
        /// <param name="part">The part</param>
        /// <param name="root">The new root element</param>
        public static void SetRootElement(this OpenXmlPart part, XElement root)
        {
            if (part == null)
                throw new ArgumentNullException("part");

            if (root == null)
                return;

            using (Stream stream = part.GetStream(FileMode.Create))
            using (StreamWriter streamWriter = new StreamWriter(stream))
            using (XmlWriter xmlWriter = XmlWriter.Create(streamWriter))
                root.WriteTo(xmlWriter);
        }
    }
}
