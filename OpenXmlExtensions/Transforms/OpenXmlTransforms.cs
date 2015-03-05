/*
 * OpenXmlTransforms.cs - Transforms for Open XML Documents
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
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Transforms
{
    /// <summary>
    /// The class represents errors that occur during transforms.
    /// </summary>
    public class OpenXmlTransformException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformException" /> class.
        /// </summary>
        public OpenXmlTransformException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformException" /> class
        /// with a specified error message.
        /// </summary>
        /// <param name="message">The error message.</param>
        public OpenXmlTransformException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformException" /> class
        /// with a specified error message and a reference to the inner exception that is
        /// the cause of this exception.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The inner exception.</param>
        public OpenXmlTransformException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// This class is the abstract base class of all Open XML transforms.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class OpenXmlTransform<TDocument> : ICloneable
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Creates a deep copy of the transform.
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
        /// and return a transformed clone. The actual transform should be implemented
        /// by overriding the <see cref="OpenXmlTransform{DocumentType}.TransformInPlace" /> method
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
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.Transform(string)" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcStringTransform<TDocument> : OpenXmlTransform<TDocument>
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
            return TransformTools.FromFlatOpcString<TDocument>(result);
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
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.Transform(XDocument)" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcDocumentTransform<TDocument> : OpenXmlTransform<TDocument>
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
            return TransformTools.FromFlatOpcDocument<TDocument>(result);
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
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.TransformInPlace" />.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    public abstract class OpenXmlPackageTransform<TDocument> : OpenXmlTransform<TDocument>
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

            using (var packageDocument = TransformTools.FromFlatOpcString<TDocument>(text))
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

            using (var packageDocument = TransformTools.FromFlatOpcDocument<TDocument>(document))
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

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on instances of <see cref="WordprocessingDocument" />.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.TransformInPlace" />.
    /// The other methods will delegate the actual transformation to this method.
    /// Currently, this class contains specific methods for transforming <see cref="Document" />,
    /// <see cref="Styles" />, and <see cref="Numbering" />. More methods can and will be added
    /// as the need arises.
    /// </remarks>
    public abstract class WordprocessingDocumentTransform : OpenXmlPackageTransform<WordprocessingDocument>
    {
        private WordprocessingDocument _template;

        #region Properties

        /// <summary>
        /// Gets or sets the template <see cref="WordprocessingDocument" />.
        /// </summary>
        [SuppressMessage("ReSharper", "NotResolvedInText")]
        public virtual WordprocessingDocument Template
        {
            get { return _template; }

            set
            {
                if (value == null)
                {
                    _template = null;
                    return;
                }

                // Check template's validity. A "minimum document" must have at least
                // a MainDocumentPart with a w:document element that has a w:body child.
                if (value.MainDocumentPart == null ||
                    value.MainDocumentPart.Document == null ||
                    value.MainDocumentPart.Document.Body == null)
                {
                    throw new ArgumentException("Illegal WordprocessingDocument", "Template");
                }

                if (value.DocumentType == WordprocessingDocumentType.Document ||
                    value.DocumentType == WordprocessingDocumentType.MacroEnabledDocument)
                {
                    _template = value;
                }
                else
                {
                    _template = (WordprocessingDocument) value.Clone();
                    _template.ChangeDocumentType(_template.DocumentType == WordprocessingDocumentType.Template
                        ? WordprocessingDocumentType.Document
                        : WordprocessingDocumentType.MacroEnabledDocument);
                    _template.Save();
                }
            }
        }

        /// <summary>
        /// Gets the template's w:body element. Returns null if no template was specified.
        /// </summary>
        protected virtual Body TemplateBody
        {
            get
            {
                return Template != null ? Template.MainDocumentPart.Document.Body : null;
            }
        }

        /// <summary>
        /// Gets the template's w:styles element. Returns null if no template was specified
        /// or there is no w:styles element.
        /// </summary>
        protected virtual Styles TemplateStyles
        {
            get
            {
                if (Template != null && Template.MainDocumentPart.StyleDefinitionsPart != null)
                    return Template.MainDocumentPart.StyleDefinitionsPart.Styles;

                return null;
            }
        }

        /// <summary>
        /// Gets the template's w:numbering element. Returns null if no template was specified
        /// or there is no w:numbering element.
        /// </summary>
        protected virtual Numbering TemplateNumbering
        {
            get
            {
                if (Template != null && Template.MainDocumentPart.NumberingDefinitionsPart != null)
                    return Template.MainDocumentPart.NumberingDefinitionsPart.Numbering;

                return null;
            }
        }

        #endregion Properties

        #region Document

        /// <summary>
        /// Replaces the root element of the <see cref="MainDocumentPart" /> contained in
        /// the given <see cref="WordprocessingDocument" /> with a transformed instance of
        /// the <see cref="Document" /> class, calling the
        /// <see cref="TransformDocument(OpenXmlElement, WordprocessingDocument)" />
        /// method to perform the actual transform.
        /// Adds a <see cref="MainDocumentPart" /> in case it does not exist, calling the
        /// <see cref="CreateDocument(WordprocessingDocument)" />
        /// method to produce the new <see cref="Document" /> element.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformDocument(WordprocessingDocument packageDocument)
        {
            var part = packageDocument.MainDocumentPart;
            if (part != null)
            {
                part.Document = (Document) TransformDocument(part.Document, packageDocument);
                part.Document.Save();
            }
            else
            {
                part = packageDocument.AddMainDocumentPart();
                part.Document = CreateDocument(packageDocument);
                part.Document.Save();
            }
            return packageDocument;
        }

        /// <summary>
        /// Transforms the <see cref="Document" /> element and its children.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="MainDocumentPart" />. The default implementation produces
        /// a deep clone of the <see cref="OpenXmlElement" />.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement" /> to be transformed.</param>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformDocument(OpenXmlElement element, WordprocessingDocument packageDocument)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Document" /> element with at least a <see cref="Body" />
        /// element (i.e., a "minimum document").
        /// </summary>
        /// <remarks>
        /// This method can be overridden by subclasses wishing to create a specific
        /// <see cref="Document" /> tree in case the <see cref="MainDocumentPart" />
        /// was previously empty.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Document" />.</returns>
        protected virtual Document CreateDocument(WordprocessingDocument packageDocument)
        {
            return new Document(new Body());
        }

        #endregion Document

        #region Styles

        /// <summary>
        /// Replaces the root element of the <see cref="StyleDefinitionsPart" /> contained
        /// in the given <see cref="WordprocessingDocument" /> with a transformed instance
        /// of the <see cref="Styles" /> class, calling the
        /// <see cref="TransformStyles(OpenXmlElement, WordprocessingDocument)" />
        /// method to perform the actual transform.
        /// Adds a <see cref="StyleDefinitionsPart" /> in case it does not exist, calling the
        /// <see cref="CreateStyles(WordprocessingDocument)" />
        /// method to produce the new <see cref="Styles" /> element.
        /// Removes the <see cref="StyleDefinitionsPart" />, or doesn't create one, if these
        /// methods return null.
        /// Also replaces the root element of the <see cref="StylesWithEffectsPart" /> with
        /// a full clone of the transformed <see cref="Styles" /> element, or removes it
        /// in case the <see cref="StyleDefinitionsPart" /> was also removed.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformStyles(WordprocessingDocument packageDocument)
        {
            // Transform the StyleDefinitionsPart's root element.
            var part = packageDocument.MainDocumentPart.StyleDefinitionsPart;
            if (part != null)
            {
                // The WordprocessingDocument has a StyleDefinitionsPart.
                // So, we transform its root element and either replace the existing
                // root element or, if the transformation results in a null element,
                // delete the StyleDefinitionsPart.
                var styles = (Styles) TransformStyles(part.Styles, packageDocument);
                if (styles != null)
                {
                    part.Styles = styles;
                    part.Styles.Save();
                }
                else
                {
                    packageDocument.MainDocumentPart.DeletePart(part);
                    part = null;
                }
            }
            else
            {
                // The WordprocessingDocument does not have a StyleDefinitionsPart.
                // If the transform produces a non-null root element, we add a new
                // StyleDefinitionsPart with that root element.
                var styles = CreateStyles(packageDocument);
                if (styles != null)
                {
                    part = packageDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    part.Styles = styles;
                    part.Styles.Save();
                }
            }

            // Make sure the StylesWithEffectPart equals the StyleDefinitionsPart.
            var effectsPart = packageDocument.MainDocumentPart.StylesWithEffectsPart;
            if (part != null)
            {
                if (effectsPart == null)
                    effectsPart = packageDocument.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                effectsPart.Styles = (Styles) part.Styles.CloneNode(true);
                effectsPart.Styles.Save();
            }
            else if (effectsPart != null)
            {
                packageDocument.MainDocumentPart.DeletePart(effectsPart);
            }

            return packageDocument;
        }

        /// <summary>
        /// Transforms the <see cref="Styles" /> element and its descendants.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="StyleDefinitionsPart" /> and its descendants. The default
        /// implementation produces a deep clone of the <see cref="OpenXmlElement" />.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement" /> to be transformed.</param>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformStyles(OpenXmlElement element, WordprocessingDocument packageDocument)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Styles" /> element as desired or returns null to not
        /// create any styles.
        /// </summary>
        /// <remarks>
        /// This method is called in case the <see cref="StyleDefinitionsPart" /> does not exist
        /// It can be overridden by subclasses wishing to create a specific <see cref="Styles" />
        /// tree. If null is returned, the <see cref="StyleDefinitionsPart" /> will not be created.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Styles" /> or null.</returns>
        protected virtual Styles CreateStyles(WordprocessingDocument packageDocument)
        {
            return null;
        }

        #endregion Styles

        #region Numbering

        /// <summary>
        /// Replaces the root element of the <see cref="NumberingDefinitionsPart" /> contained
        /// in the given <see cref="WordprocessingDocument" /> with a transformed instance
        /// of the <see cref="Numbering" /> class, calling the
        /// <see cref="TransformNumbering(OpenXmlElement, WordprocessingDocument)" />
        /// method to perform the actual transform.
        /// Adds a <see cref="NumberingDefinitionsPart" /> in case it does not exist, calling the
        /// <see cref="CreateNumbering(WordprocessingDocument)" />
        /// method to produce the new <see cref="Numbering" /> element.
        /// Removes the <see cref="NumberingDefinitionsPart" />, or doesn't create one, if these
        /// methods return null.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformNumbering(WordprocessingDocument packageDocument)
        {
            var part = packageDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (part != null)
            {
                var numbering = (Numbering) TransformNumbering(part.Numbering, packageDocument);
                if (numbering != null)
                {
                    part.Numbering = numbering;
                    part.Numbering.Save();
                }
                else
                {
                    packageDocument.MainDocumentPart.DeletePart(part);
                }
            }
            else
            {
                var numbering = CreateNumbering(packageDocument);
                if (numbering != null)
                {
                    part = packageDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                    part.Numbering = numbering;
                    part.Numbering.Save();
                }
            }
            return packageDocument;
        }

        /// <summary>
        /// Transforms the <see cref="Numbering" /> element and its descendants.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="NumberingDefinitionsPart" /> and its descendants. The default
        /// implementation just produces a deep clone of the <see cref="OpenXmlElement" />.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement" /> to be transformed.</param>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformNumbering(OpenXmlElement element, WordprocessingDocument packageDocument)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Numbering" /> element as desired or returns null to not
        /// create any numbering definitions.
        /// </summary>
        /// <remarks>
        /// This method is called in case the <see cref="NumberingDefinitionsPart" /> does not exist
        /// It can be overridden by subclasses wishing to create a specific <see cref="Numbering" />
        /// tree. If null is returned, the <see cref="NumberingDefinitionsPart" /> will not be created.
        /// </remarks>
        /// <param name="packageDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Numbering" /> or null.</returns>
        protected virtual Numbering CreateNumbering(WordprocessingDocument packageDocument)
        {
            return null;
        }

        #endregion Numbering

        #region Headers

        protected WordprocessingDocument TransformHeaders(WordprocessingDocument packageDocument)
        {
            var parts = packageDocument.MainDocumentPart.HeaderParts.ToList();
            foreach (var part in parts)
            {
                var header = (Header) TransformHeader(part.Header, packageDocument);
                if (header != null)
                {
                    part.Header = header;
                    part.Header.Save();
                }
                else
                {
                    packageDocument.MainDocumentPart.DeletePart(part);
                }
            }
            return packageDocument;
        }

        protected virtual object TransformHeader(OpenXmlElement element, WordprocessingDocument packageDocument)
        {
            return element.CloneNode(true);
        }

        #endregion

        #region Footers

        protected WordprocessingDocument TransformFooters(WordprocessingDocument packageDocument)
        {
            var parts = packageDocument.MainDocumentPart.FooterParts.ToList();
            foreach (var part in parts)
            {
                var footer = (Footer) TransformFooter(part.Footer, packageDocument);
                if (footer != null)
                {
                    part.Footer = footer;
                    part.Footer.Save();
                }
                else
                {
                    packageDocument.MainDocumentPart.DeletePart(part);
                }
            }
            return packageDocument;
        }

        protected virtual object TransformFooter(OpenXmlElement element, WordprocessingDocument packageDocument)
        {
            return element.CloneNode(true);
        }

        #endregion

        #region ICloneable Methods

        /// <summary>
        /// Creates a deep copy of the transform.
        /// </summary>
        /// <returns>The clone.</returns>
        public override object Clone()
        {
            var transform = (WordprocessingDocumentTransform)base.Clone();
            if (Template != null)
                transform.Template = (WordprocessingDocument)Template.Clone();

            return transform;
        }

        #endregion
    }
}
