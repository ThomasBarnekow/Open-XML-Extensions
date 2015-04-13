/*
 * TransformationTools.cs - Tools for Transformations
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

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ContractArchitect.OpenXml.Transformation
{
    /// <summary>
    /// Utility class providing a number of tools for transforms.
    /// </summary>
    public static class TransformationTools
    {
        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC <see cref="XDocument" />.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument" />.</param>
        /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static TDocument FromFlatOpcDocument<TDocument>(XDocument document)
            where TDocument : OpenXmlPackage
        {
            var t = typeof (TDocument);
            if (t == typeof (WordprocessingDocument))
            {
                return WordprocessingDocument.FromFlatOpcDocument(document) as TDocument;
            }
            if (t == typeof (SpreadsheetDocument))
            {
                return SpreadsheetDocument.FromFlatOpcDocument(document) as TDocument;
            }
            if (t == typeof (PresentationDocument))
            {
                return PresentationDocument.FromFlatOpcDocument(document) as TDocument;
            }
            throw new OpenXmlTransformationException("Unsupported document type: " + t);
        }

        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string.</param>
        /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static TDocument FromFlatOpcString<TDocument>(string text)
            where TDocument : OpenXmlPackage
        {
            var t = typeof (TDocument);
            if (t == typeof (WordprocessingDocument))
            {
                return WordprocessingDocument.FromFlatOpcString(text) as TDocument;
            }
            if (t == typeof (SpreadsheetDocument))
            {
                return SpreadsheetDocument.FromFlatOpcString(text) as TDocument;
            }
            if (t == typeof (PresentationDocument))
            {
                return PresentationDocument.FromFlatOpcString(text) as TDocument;
            }
            throw new OpenXmlTransformationException("Unsupported document type: " + t);
        }
    }
}
