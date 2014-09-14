﻿/*
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
using System.Collections.Generic;
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
        /// Initializes a new instance of the <see cref="OpenXmlTransformException"/> class.
        /// </summary>
        public OpenXmlTransformException()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformException"/> class
        /// with a specified error message.
        /// </summary>
        /// <param name="message">The error message.</param>
        public OpenXmlTransformException(string message)
            : base(message)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformException"/> class
        /// with a specified error message and a reference to the inner exception that is 
        /// the cause of this exception. 
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The inner exception.</param>
        public OpenXmlTransformException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }

    /// <summary>
    /// This class is the abstract base class of all Open XML transforms.
    /// </summary>
    /// <typeparam name="DocumentType">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
    public abstract class OpenXmlTransform<DocumentType> 
        where DocumentType : OpenXmlPackage
    {
        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string.</param>
        /// <returns>A new instance of DocumentType</returns>
        protected static DocumentType FromFlatOpcString(string text)
        {
            Type t = typeof(DocumentType);
            if (t == typeof(WordprocessingDocument))
                return WordprocessingDocument.FromFlatOpcString(text) as DocumentType;
            else if (t == typeof(SpreadsheetDocument))
                return SpreadsheetDocument.FromFlatOpcString(text) as DocumentType;
            else if (t == typeof(PresentationDocument))
                return PresentationDocument.FromFlatOpcString(text) as DocumentType;
            else
                throw new OpenXmlTransformException("Unsupported document type: " + t);
        }

        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/>.</param>
        /// <returns>A new instance of DocumentType</returns>
        protected static DocumentType FromFlatOpcDocument(XDocument document)
        {
            Type t = typeof(DocumentType);
            if (t == typeof(WordprocessingDocument))
                return WordprocessingDocument.FromFlatOpcDocument(document) as DocumentType;
            else if (t == typeof(SpreadsheetDocument))
                return SpreadsheetDocument.FromFlatOpcDocument(document) as DocumentType;
            else if (t == typeof(PresentationDocument))
                return PresentationDocument.FromFlatOpcDocument(document) as DocumentType;
            else
                throw new OpenXmlTransformException("Unsupported document type: " + t);
        }

        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="xml">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public virtual string Transform(string xml)
        {
            return xml;
        }

        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument"/>.</returns>
        public virtual XDocument Transform(XDocument document)
        {
            return document;
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/>.
        /// </summary>
        /// <remarks>
        /// This method, if overridden by a subclass, must clone the original document
        /// and return a transformed clone. The actual transform should be implemented
        /// by overriding the <see cref="OpenXmlTransform{DocumentType}.TransformInPlace"/> method
        /// which is called by the default implementation in this class.
        /// </remarks>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public virtual DocumentType Transform(DocumentType packageDocument)
        {
            if (packageDocument == null)
                return null;

            return TransformInPlace((DocumentType)packageDocument.Clone());
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/> in-place.
        /// </summary>
        /// <remarks>
        /// This method, if overridden by a subclass, must transform the original document
        /// in-place rather than transforming a clone. Otherwise, if called directly, it 
        /// will not have the desired effect.
        /// </remarks>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public virtual DocumentType TransformInPlace(DocumentType packageDocument)
        {
            return packageDocument;
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on a Flat OPC string.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.Transform(string)"/>.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="DocumentType">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
    public abstract class FlatOpcStringTransform<DocumentType> : OpenXmlTransform<DocumentType>
        where DocumentType : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument"/>.</returns>
        public sealed override XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            string result = Transform(document.ToString());
            return XDocument.Parse(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/>.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public sealed override DocumentType Transform(DocumentType packageDocument)
        {
            if (packageDocument == null)
                return null;

            string result = Transform(packageDocument.ToFlatOpcString());
            return FromFlatOpcString(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/> in-place.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public sealed override DocumentType TransformInPlace(DocumentType packageDocument)
        {
            if (packageDocument == null)
                throw new ArgumentNullException("packageDocument");

            return (DocumentType)packageDocument.ReplaceWith(Transform(packageDocument));
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on a Flat OPC <see cref="XDocument"/>.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.Transform(XDocument)"/>.
    /// The other methods will delegate the actual transformation to this method.
    /// </remarks>
    /// <typeparam name="DocumentType">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
    public abstract class FlatOpcDocumentTransform<DocumentType> : OpenXmlTransform<DocumentType>
        where DocumentType : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="xml">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public sealed override string Transform(string xml)
        {
            if (xml == null)
                return null;

            XDocument result = Transform(XDocument.Parse(xml));
            return result.ToString();
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/>.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The cloned and transformed document.</returns>
        public sealed override DocumentType Transform(DocumentType packageDocument)
        {
            if (packageDocument == null)
                return null;

            XDocument result = Transform(packageDocument.ToFlatOpcDocument());
            return FromFlatOpcDocument(result);
        }

        /// <summary>
        /// Transforms an instance of a subclass of <see cref="OpenXmlPackage"/> in-place.
        /// </summary>
        /// <param name="packageDocument">The document to be transformed.</param>
        /// <returns>The document (transformed in-place).</returns>
        public sealed override DocumentType TransformInPlace(DocumentType packageDocument)
        {
            if (packageDocument == null)
                throw new ArgumentNullException("packageDocument");

            return (DocumentType)packageDocument.ReplaceWith(Transform(packageDocument));
        }
    }

    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on instances of <see cref="WordprocessingDocument"/>.
    /// Derived classes are supposed to override the 
    /// 
    ///     <see cref="OpenXmlTransform{DocumentType}.TransformInPlace"/> 
    ///     
    /// method.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransform{DocumentType}.TransformInPlace"/>.
    /// The other methods will delegate the actual transformation to this method.
    /// Currently, this class contains specific methods for transforming <see cref="Document"/>,
    /// <see cref="Styles"/>, and <see cref="Numbering"/>. More methods can and will be added
    /// as the need arises.
    /// </remarks>
    public abstract class WordprocessingDocumentTransform : OpenXmlTransform<WordprocessingDocument>
    {
        /// <summary>
        /// Initializes a new instance of <see cref="WordprocessingDocumentTransform"/>.
        /// </summary>
        protected WordprocessingDocumentTransform()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of <see cref="WordprocessingDocumentTransform"/>
        /// with a template <see cref="WordprocessingDocument"/>.
        /// </summary>
        /// <param name="template">The template to be used.</param>
        protected WordprocessingDocumentTransform(WordprocessingDocument template)
            : this()
        {
            Template = template;
        }

        /// <summary>
        /// Gets or sets the template <see cref="WordprocessingDocument"/>.
        /// </summary>
        public virtual WordprocessingDocument Template { get; set; }

        /// <summary>
        /// Transforms a Flat OPC string.
        /// </summary>
        /// <param name="xml">The Flat OPC string to be transformed.</param>
        /// <returns>The transformed Flat OPC string.</returns>
        public sealed override string Transform(string xml)
        {
            if (xml == null)
                return null;

            using (WordprocessingDocument document = WordprocessingDocument.FromFlatOpcString(xml))
                return TransformInPlace(document).ToFlatOpcString();
        }

        /// <summary>
        /// Transforms a Flat OPC <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/> to be transformed.</param>
        /// <returns>The transformed Flat OPC <see cref="XDocument"/>.</returns>
        public sealed override XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.FromFlatOpcDocument(document))
                return TransformInPlace(wordDoc).ToFlatOpcDocument();
        }

        #region Document

        /// <summary>
        /// Replaces the root element of the <see cref="MainDocumentPart"/> contained in
        /// the given <see cref="WordprocessingDocument"/> with a transformed instance of
        /// the <see cref="Document"/> class, calling the
        /// 
        ///     <see cref="TransformDocument(OpenXmlElement, WordprocessingDocument)"/>
        /// 
        /// method to perform the actual transform.
        /// Adds a <see cref="MainDocumentPart"/> in case it does not exist, calling the 
        /// 
        ///     <see cref="CreateDocument(WordprocessingDocument)"/> 
        /// 
        /// method to produce the new <see cref="Document"/> element.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// 
        ///     <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace"/>
        ///     
        /// method.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument"/>.</returns>
        protected WordprocessingDocument TransformDocument(WordprocessingDocument wordDoc)
        {
            MainDocumentPart part = wordDoc.MainDocumentPart;
            if (part != null)
            {
                part.Document = (Document)TransformDocument(part.Document, wordDoc);
            }
            else
            {
                part = wordDoc.AddMainDocumentPart();
                part.Document = CreateDocument(wordDoc);
            }
            return wordDoc;
        }

        /// <summary>
        /// Transforms the <see cref="Document"/> element and its children.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="MainDocumentPart"/>. The default implementation produces
        /// a deep clone of the <see cref="OpenXmlElement"/>.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement"/> to be transformed.</param>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="OpenXmlElement"/>.</returns>
        protected virtual object TransformDocument(OpenXmlElement element, WordprocessingDocument wordDoc)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Document"/> element with at least a <see cref="Body"/>
        /// element (i.e., a "minimum document"). 
        /// </summary>
        /// <remarks>
        /// This method can be overridden by subclasses wishing to create a specific
        /// <see cref="Document"/> tree in case the <see cref="MainDocumentPart"/>
        /// was previously empty.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed.</param>
        /// <returns>A new instance of <see cref="Document"/>.</returns>
        protected virtual Document CreateDocument(WordprocessingDocument wordDoc)
        {
            return new Document(new Body());
        }

        #endregion Document

        #region Styles

        /// <summary>
        /// Replaces the root element of the <see cref="StyleDefinitionsPart"/> contained
        /// in the given <see cref="WordprocessingDocument"/> with a transformed instance
        /// of the <see cref="Styles"/> class, calling the
        /// 
        ///     <see cref="TransformStyles(OpenXmlElement, WordprocessingDocument)"/>
        ///     
        /// method to perform the actual transform.
        /// Adds a <see cref="StyleDefinitionsPart"/> in case it does not exist, calling the 
        /// 
        ///     <see cref="CreateStyles(WordprocessingDocument)"/> 
        /// 
        /// method to produce the new <see cref="Styles"/> element.
        /// Removes the <see cref="StyleDefinitionsPart"/>, or doesn't create one, if these
        /// methods return null.
        /// 
        /// Also replaces the root element of the <see cref="StylesWithEffectsPart"/> with 
        /// a full clone of the transformed <see cref="Styles"/> element, or removes it
        /// in case the <see cref="StyleDefinitionsPart"/> was also removed.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// 
        ///     <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace"/>
        ///     
        /// method.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument"/>.</returns>
        protected WordprocessingDocument TransformStyles(WordprocessingDocument wordDoc)
        {
            // Transform the StyleDefinitionsPart's root element.
            StyleDefinitionsPart part = wordDoc.MainDocumentPart.StyleDefinitionsPart;
            if (part != null)
            {
                // The WordprocessingDocument has a StyleDefinitionsPart.
                // So, we transform its root element and either replace the existing
                // root element or, if the transformation results in a null element,
                // delete the StyleDefinitionsPart.
                Styles styles = (Styles)TransformStyles(part.Styles, wordDoc);
                if (styles != null)
                {
                    part.Styles = styles;
                }
                else
                {
                    wordDoc.MainDocumentPart.DeletePart(part);
                    part = null;
                }
            }
            else
            {
                // The WordprocessingDocument does not have a StyleDefinitionsPart.
                // If the transform produces a non-null root element, we add a new
                // StyleDefinitionsPart with that root element.
                Styles styles = CreateStyles(wordDoc);
                if (styles != null)
                {
                    part = wordDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    part.Styles = styles;
                }
            }

            // Make sure the StylesWithEffectPart equals the StyleDefinitionsPart.
            StylesWithEffectsPart effectsPart = wordDoc.MainDocumentPart.StylesWithEffectsPart;
            if (part != null)
            {
                if (effectsPart == null)
                    effectsPart = wordDoc.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                effectsPart.Styles = (Styles)part.Styles.CloneNode(true);
            }
            else if (effectsPart != null)
            {
                wordDoc.MainDocumentPart.DeletePart(effectsPart);
            }

            return wordDoc;
        }

        /// <summary>
        /// Transforms the <see cref="Styles"/> element and its descendants.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="StyleDefinitionsPart"/> and its descendants. The default 
        /// implementation produces a deep clone of the <see cref="OpenXmlElement"/>.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement"/> to be transformed.</param>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement"/>.</returns>
        protected virtual object TransformStyles(OpenXmlElement element, WordprocessingDocument wordDoc)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Styles"/> element as desired or returns null to not
        /// create any styles.
        /// </summary>
        /// <remarks>
        /// This method is called in case the <see cref="StyleDefinitionsPart"/> does not exist
        /// It can be overridden by subclasses wishing to create a specific <see cref="Styles"/> 
        /// tree. If null is returned, the <see cref="StyleDefinitionsPart"/> will not be created.
        /// This is the default.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed.</param>
        /// <returns>A new instance of <see cref="Styles"/> or null.</returns>
        protected virtual Styles CreateStyles(WordprocessingDocument wordDoc)
        {
            return null;
        }

        #endregion Styles

        #region Numbering

        /// <summary>
        /// Replaces the root element of the <see cref="NumberingDefinitionsPart"/> contained
        /// in the given <see cref="WordprocessingDocument"/> with a transformed instance
        /// of the <see cref="Numbering"/> class, calling the
        /// 
        ///     <see cref="TransformNumbering(OpenXmlElement, WordprocessingDocument)"/>
        ///     
        /// method to perform the actual transform.
        /// Adds a <see cref="NumberingDefinitionsPart"/> in case it does not exist, calling the 
        /// 
        ///     <see cref="CreateNumbering(WordprocessingDocument)"/> 
        /// 
        /// method to produce the new <see cref="Numbering"/> element.
        /// Removes the <see cref="NumberingDefinitionsPart"/>, or doesn't create one, if these
        /// methods return null.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// 
        ///     <see cref="OpenXmlTransform{WordprocessingDocument}.TransformInPlace"/>
        ///     
        /// method.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument"/>.</returns>
        protected WordprocessingDocument TransformNumbering(WordprocessingDocument wordDoc)
        {
            NumberingDefinitionsPart part = wordDoc.MainDocumentPart.NumberingDefinitionsPart;
            if (part != null)
            {
                Numbering numbering = (Numbering)TransformNumbering(part.Numbering, wordDoc);
                if (numbering != null)
                    part.Numbering = numbering;
                else
                    wordDoc.MainDocumentPart.DeletePart(part);
            }
            else
            {
                Numbering numbering = CreateNumbering(wordDoc);
                if (numbering != null)
                {
                    part = wordDoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                    part.Numbering = numbering;
                }
            }
            return wordDoc;
        }

        /// <summary>
        /// Transforms the <see cref="Numbering"/> element and its descendants.
        /// </summary>
        /// <remarks>
        /// This method will be overridden by subclasses wishing to transform the root element
        /// of the <see cref="NumberingDefinitionsPart"/> and its descendants. The default 
        /// implementation just produces a deep clone of the <see cref="OpenXmlElement"/>.
        /// </remarks>
        /// <param name="element">The <see cref="OpenXmlElement"/> to be transformed.</param>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement"/>.</returns>
        protected virtual object TransformNumbering(OpenXmlElement element, WordprocessingDocument wordDoc)
        {
            return element.CloneNode(true);
        }

        /// <summary>
        /// Creates a new <see cref="Numbering"/> element as desired or returns null to not
        /// create any numbering definitions.
        /// </summary>
        /// <remarks>
        /// This method is called in case the <see cref="NumberingDefinitionsPart"/> does not exist
        /// It can be overridden by subclasses wishing to create a specific <see cref="Numbering"/> 
        /// tree. If null is returned, the <see cref="NumberingDefinitionsPart"/> will not be created.
        /// This is the default.
        /// </remarks>
        /// <param name="wordDoc">The <see cref="WordprocessingDocument"/> to be transformed.</param>
        /// <returns>A new instance of <see cref="Numbering"/> or null.</returns>
        protected virtual Numbering CreateNumbering(WordprocessingDocument wordDoc)
        {
            return null;
        }

        #endregion Numbering
    }
}
