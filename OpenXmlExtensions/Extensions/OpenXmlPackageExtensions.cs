/*
 * OpenXmlPackageExtensions.cs - Extensions for OpenXmlPackage
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
using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    /// Extensions for <see cref="OpenXmlPackage" />.
    /// </summary>
    public static class OpenXmlPackageExtensions
    {
        /// <summary>
        /// Gets all parts contained in the <see cref="OpenXmlPackage" /> in a
        /// breadth-first fashion, i.e., the direct and indirect relationship
        /// targets of the package (where the <see cref="OpenXmlPartContainer.Parts" />
        /// property only returns the direct relationship targets).
        /// </summary>
        public static IEnumerable<OpenXmlPart> GetAllParts(this OpenXmlPackage package)
        {
            return new OpenXmlParts(package);
        }

        /// <summary>
        /// Replaces the document's contents with the contents of the given replacement's contents.
        /// </summary>
        /// <param name="document">The destination document</param>
        /// <param name="replacement">The source document</param>
        /// <returns>The original document with replaced contents</returns>
        public static OpenXmlPackage ReplaceWith(this OpenXmlPackage document,
            OpenXmlPackage replacement)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (replacement == null)
                throw new ArgumentNullException("replacement");

            // Delete all parts (i.e., the direct relationship targets and their
            // children).
            document.DeleteParts(document.GetPartsOfType<OpenXmlPart>());

            // Add the replacement's parts to the document.
            foreach (var part in replacement.Parts)
                document.AddPart(part.OpenXmlPart, part.RelationshipId);

            // Save and return.
            document.Package.Flush();
            return document;
        }
    }

    /// <summary>
    /// Enumeration of all parts contained in an <see cref="OpenXmlPackage" />
    /// rather than just the direct relationship targets.
    /// </summary>
    public class OpenXmlParts : IEnumerable<OpenXmlPart>
    {
        private readonly OpenXmlPackage _package;

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the OpenXmlPackagePartIterator class using the supplied OpenXmlPackage class.
        /// </summary>
        /// <param name="package">The OpenXmlPackage to use to enumerate parts</param>
        public OpenXmlParts(OpenXmlPackage package)
        {
            _package = package;
        }

        #endregion

        #region IEnumerable<OpenXmlPart> Members

        /// <summary>
        /// Gets an enumerator for parts in the whole package.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<OpenXmlPart> GetEnumerator()
        {
            var parts = new List<OpenXmlPart>();
            var queue = new Queue<OpenXmlPart>();

            // Enqueue all direct relationship targets.
            foreach (var target in _package.Parts)
            {
                queue.Enqueue(target.OpenXmlPart);
            }

            while (queue.Count > 0)
            {
                // Add next part from queue to the set of parts to be returned.
                var part = queue.Dequeue();
                parts.Add(part);

                // Enqueue all direct relationship targets of current part that
                // are not already enqueued or in the set of parts to be returned.
                foreach (var indirectTarget in part.Parts)
                {
                    if (!queue.Contains(indirectTarget.OpenXmlPart) &&
                        !parts.Contains(indirectTarget.OpenXmlPart))
                    {
                        queue.Enqueue(indirectTarget.OpenXmlPart);
                    }
                }
            }

            // Done.
            return parts.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator for parts in the whole package.
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion
    }
}
