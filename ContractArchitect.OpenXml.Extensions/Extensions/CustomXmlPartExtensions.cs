/*
 * CustomXmlPartExtensions.cs - Extensions for CustomXmlPart
 * 
 * Copyright 2015 Thomas Barnekow
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
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ContractArchitect.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for the <see cref="CustomXmlPart" /> class.
    /// </summary>
    public static class CustomXmlPartExtensions
    {
        public static CustomXmlPart ReplaceNamespace(this CustomXmlPart part, XNamespace newNs, string newPrefix)
        {
            if (part == null)
                throw new ArgumentNullException("part");
            if (newNs == null)
                throw new ArgumentNullException("newNs");

            var oldRoot = part.GetRootElement();
            if (oldRoot == null)
                return null;

            var oldNs = oldRoot.Name.Namespace;
            var newRoot = ReplaceNamespace(oldRoot, oldNs, newNs, newPrefix);
            part.SetRootElement(newRoot);

            return part;
        }

        public static XElement ReplaceNamespace(XElement element, XNamespace oldNs, XNamespace newNs, string newPrefix)
        {
            return (XElement) TransformNamespace(element, oldNs, newNs, newPrefix);
        }

        internal static XObject TransformNamespace(XObject obj, XNamespace oldNs, XNamespace newNs, string newPrefix)
        {
            var element = obj as XElement;
            if (element != null)
            {
                if (element.Name.Namespace == oldNs)
                {
                    return new XElement(newNs.GetName(element.Name.LocalName),
                        element.Attributes().Select(a => TransformNamespace(a, oldNs, newNs, newPrefix)),
                        element.Nodes().Select(n => TransformNamespace(n, oldNs, newNs, newPrefix)));
                }
                return new XElement(element.Name,
                    element.Attributes().Select(a => TransformNamespace(a, oldNs, newNs, newPrefix)),
                    element.Nodes().Select(n => TransformNamespace(n, oldNs, newNs, newPrefix)));
            }

            var attribute = obj as XAttribute;
            if (attribute != null)
            {
                if (attribute.Name.Namespace == oldNs)
                {
                    return new XAttribute(newNs.GetName(attribute.Name.LocalName), attribute.Value);
                }
                if (attribute.Name.Namespace == XNamespace.Xmlns && attribute.Value == oldNs.NamespaceName)
                {
                    return new XAttribute(XNamespace.Xmlns + newPrefix, newNs.NamespaceName);
                }
                return attribute;
            }

            return obj;
        }
    }
}
