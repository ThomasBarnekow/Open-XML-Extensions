/*
 * OpenXmlTransformations.cs - Transformations for Open XML Documents
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
using System.Xml.Linq;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transformation
{
    /// <summary>
    /// This class is the abstract base class of all Open XML transformations.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class OpenXmlTransformation<TDocument> : ICloneable
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Creates a deep copy of the transformation.
        /// </summary>
        /// <returns></returns>
        public virtual object Clone()
        {
            return MemberwiseClone();
        }

        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public virtual string Transform(string text)
        {
            return text;
        }

        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument" />.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument" /> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument" />.</returns>
        public virtual XDocument Transform(XDocument document)
        {
            return document;
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" />.
        /// </summary>
        /// <remarks>
        /// This method, if overridden by a subclass, must clone the original document
        /// and return a transformed clone. The actual transformation should be implemented
        /// by overriding the <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" /> method
        /// which is called by the default implementation in this class.
        /// </remarks>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public virtual TDocument Transform(TDocument packageDocument)
        {
            return packageDocument == null ? null : TransformInPlace((TDocument) packageDocument.Clone());
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" /> in-place.
        /// </summary>
        /// <remarks>
        /// This method, if overridden by a subclass, must transform the original document
        /// in-place rather than transforming a clone. Otherwise, if called directly, it
        /// will not have the desired effect.
        /// </remarks>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public virtual TDocument TransformInPlace(TDocument packageDocument)
        {
            return packageDocument;
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on a Flat OPC string.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransformation{TDocument}.Transform(string)" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcStringTransformation<TDocument> : OpenXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument" />.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument" /> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument" />.</returns>
        public override sealed XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            var result = Transform(document.ToString());
            return XDocument.Parse(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" />.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public override sealed TDocument Transform(TDocument packageDocument)
        {
            if (packageDocument == null)
                return null;

            var result = Transform(packageDocument.ToFlatOpcString());
            return TransformationTools.FromFlatOpcString<TDocument>(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" /> in-place.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public override sealed TDocument TransformInPlace(TDocument packageDocument)
        {
            if (packageDocument == null)
                throw new ArgumentNullException("packageDocument");

            return (TDocument) packageDocument.ReplaceWith(Transform(packageDocument));
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on a Flat OPC <see cref="XDocument" />.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransformation{TDocument}.Transform(XDocument)" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcDocumentTransformation<TDocument> : OpenXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public override sealed string Transform(string text)
        {
            if (text == null)
                return null;

            var result = Transform(XDocument.Parse(text));
            return result.ToString();
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" />.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public override sealed TDocument Transform(TDocument packageDocument)
        {
            if (packageDocument == null)
                return null;

            var result = Transform(packageDocument.ToFlatOpcDocument());
            return TransformationTools.FromFlatOpcDocument<TDocument>(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage" /> in-place.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public override sealed TDocument TransformInPlace(TDocument packageDocument)
        {
            if (packageDocument == null)
                throw new ArgumentNullException("packageDocument");

            return (TDocument) packageDocument.ReplaceWith(Transform(packageDocument));
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on instances of <see cref="OpenXmlPackage" /> or, more specifically,
    /// instances of its subclasses.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    public abstract class OpenXmlPackageTransformation<TDocument> : OpenXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public override sealed string Transform(string text)
        {
            if (text == null)
                return null;

            using (var packageDocument = TransformationTools.FromFlatOpcString<TDocument>(text))
                return TransformInPlace(packageDocument).ToFlatOpcString();
        }

        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument" />.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument" /> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument" />.</returns>
        public override sealed XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            using (var packageDocument = TransformationTools.FromFlatOpcDocument<TDocument>(document))
                return TransformInPlace(packageDocument).ToFlatOpcDocument();
        }

        /// <summary>
        /// Transforms an Open XML package document.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public override sealed TDocument Transform(TDocument packageDocument)
        {
            return packageDocument == null ? null : TransformInPlace((TDocument) packageDocument.Clone());
        }
    }
}
