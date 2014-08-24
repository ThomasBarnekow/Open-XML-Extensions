/*
 * OpenXmlTransformExtensions.cs - Extension Methods for Open XML Transforms
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
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transforms.Extensions
{
    /// <summary>
    /// Extension methods used by OpenXmlTransforms.
    /// </summary>
    public static class OpenXmlTransformExtensions
    {
        /// <summary>
        /// Copies the document. The copy will be backed by a <see cref="MemoryStream"/>.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The copy</returns>
        public static WordprocessingDocument Copy(this WordprocessingDocument document)
        {
            return Copy(document, new MemoryStream());
        }

        /// <summary>
        /// Copies the document. The copy will be backed by the given <see cref="Stream"/>.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="stream"></param>
        /// <returns>The copy</returns>
        public static WordprocessingDocument Copy(this WordprocessingDocument document, Stream stream)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (stream == null)
                throw new ArgumentNullException("stream");

            // Create new WordprocessingDocument backed by stream.
            WordprocessingDocument copy = WordprocessingDocument.Create(stream, document.DocumentType);

            // Copy all document parts.
            foreach (var part in document.Parts)
                copy.AddPart(part.OpenXmlPart, part.RelationshipId);

            copy.Package.Flush();
            return copy;
        }

        /// <summary>
        /// Replaces the document's contents with the contents of the given replacement's contents.
        /// </summary>
        /// <param name="document">The destination document</param>
        /// <param name="replacement">The source document</param>
        /// <returns>The original document with replaced contents</returns>
        public static WordprocessingDocument ReplaceWith(this WordprocessingDocument document, WordprocessingDocument replacement)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (replacement == null)
                throw new ArgumentNullException("replacement");

            // Delete all parts (every part is an OpenXmlPart).
            document.DeleteParts(document.GetPartsOfType<OpenXmlPart>());

            // Add the replacement's parts to the document.
            foreach (var part in replacement.Parts)
                document.AddPart(part.OpenXmlPart, part.RelationshipId);

            // Save and return.
            document.Package.Flush();
            return document;
        }
    }
}
