/*
 * WordprocessingDocumentTransformation.cs - Transformations for WordprocessingDocuments
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
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Transformation
{
    /// <summary>
    /// This class should be subclassed by concrete transforms that perform their specific
    /// transformation task on instances of <see cref="WordprocessingDocument" />.
    /// </summary>
    /// <remarks>
    /// Subclasses must override <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" />.
    /// The other methods will delegate the actual transformation to this method.
    /// Currently, this class contains specific methods for transforming <see cref="Document" />,
    /// <see cref="Styles" />, and <see cref="Numbering" />. More methods can and will be added
    /// as the need arises.
    /// </remarks>
    public abstract class WordprocessingDocumentTransformation : OpenXmlPackageTransformation<WordprocessingDocument>
    {
        private WordprocessingDocument _template;

        #region Properties

        /// <summary>
        /// Gets or sets the template <see cref="WordprocessingDocument" />.
        /// </summary>
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
                    throw new ArgumentException("Illegal WordprocessingDocument", "value");
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
        protected Body TemplateBody
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
        protected Styles TemplateStyles
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
        protected Numbering TemplateNumbering
        {
            get
            {
                if (Template != null && Template.MainDocumentPart.NumberingDefinitionsPart != null)
                    return Template.MainDocumentPart.NumberingDefinitionsPart.Numbering;

                return null;
            }
        }

        #endregion Properties

        public override WordprocessingDocument TransformInPlace(WordprocessingDocument wordDocument)
        {
            return base.TransformInPlace(wordDocument);
        }

        #region Document

        /// <summary>
        /// Replaces the root element of the <see cref="MainDocumentPart" /> contained in
        /// the given <see cref="WordprocessingDocument" /> with a transformed instance of
        /// the <see cref="Document" /> class, calling the
        /// <see cref="TransformDocument(OpenXmlElement, WordprocessingDocument)" />
        /// method to perform the actual transformation.
        /// Adds a <see cref="MainDocumentPart" /> in case it does not exist, calling the
        /// <see cref="CreateDocument(WordprocessingDocument)" />
        /// method to produce the new <see cref="Document" /> element.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformDocument(WordprocessingDocument wordDocument)
        {
            var part = wordDocument.MainDocumentPart;
            if (part != null)
            {
                part.Document = (Document) TransformDocument(part.Document, wordDocument);
                //part.Document.Save();
            }
            else
            {
                part = wordDocument.AddMainDocumentPart();
                part.Document = CreateDocument(wordDocument);
                //part.Document.Save();
            }
            return wordDocument;
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformDocument(OpenXmlElement element, WordprocessingDocument wordDocument)
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Document" />.</returns>
        protected virtual Document CreateDocument(WordprocessingDocument wordDocument)
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
        /// method to perform the actual transformation.
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
        /// <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformStyles(WordprocessingDocument wordDocument)
        {
            // Transform the StyleDefinitionsPart's root element.
            var part = wordDocument.MainDocumentPart.StyleDefinitionsPart;
            if (part != null)
            {
                // The WordprocessingDocument has a StyleDefinitionsPart.
                // So, we transform its root element and either replace the existing
                // root element or, if the transformation results in a null element,
                // delete the StyleDefinitionsPart.
                var styles = (Styles) TransformStyles(part.Styles, wordDocument);
                if (styles != null)
                {
                    part.Styles = styles;
                    //part.Styles.Save();
                }
                else
                {
                    wordDocument.MainDocumentPart.DeletePart(part);
                    part = null;
                }
            }
            else
            {
                // The WordprocessingDocument does not have a StyleDefinitionsPart.
                // If the transformation produces a non-null root element, we add a new
                // StyleDefinitionsPart with that root element.
                var styles = CreateStyles(wordDocument);
                if (styles != null)
                {
                    part = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    part.Styles = styles;
                    //part.Styles.Save();
                }
            }

            // Make sure the StylesWithEffectPart equals the StyleDefinitionsPart.
            var effectsPart = wordDocument.MainDocumentPart.StylesWithEffectsPart;
            if (part != null)
            {
                if (effectsPart == null)
                    effectsPart = wordDocument.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                effectsPart.Styles = (Styles) part.Styles.CloneNode(true);
                //effectsPart.Styles.Save();
            }
            else if (effectsPart != null)
            {
                wordDocument.MainDocumentPart.DeletePart(effectsPart);
            }

            return wordDocument;
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformStyles(OpenXmlElement element, WordprocessingDocument wordDocument)
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Styles" /> or null.</returns>
        protected virtual Styles CreateStyles(WordprocessingDocument wordDocument)
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
        /// method to perform the actual transformation.
        /// Adds a <see cref="NumberingDefinitionsPart" /> in case it does not exist, calling the
        /// <see cref="CreateNumbering(WordprocessingDocument)" />
        /// method to produce the new <see cref="Numbering" /> element.
        /// Removes the <see cref="NumberingDefinitionsPart" />, or doesn't create one, if these
        /// methods return null.
        /// </summary>
        /// <remarks>
        /// This method is meant to be called from overrides of the
        /// <see cref="OpenXmlTransformation{TDocument}.TransformInPlace" />
        /// method.
        /// </remarks>
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed in-place.</param>
        /// <returns>The transformed <see cref="WordprocessingDocument" />.</returns>
        protected WordprocessingDocument TransformNumbering(WordprocessingDocument wordDocument)
        {
            var part = wordDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (part != null)
            {
                var numbering = (Numbering) TransformNumbering(part.Numbering, wordDocument);
                if (numbering != null)
                {
                    part.Numbering = numbering;
                    //part.Numbering.Save();
                }
                else
                {
                    wordDocument.MainDocumentPart.DeletePart(part);
                }
            }
            else
            {
                var numbering = CreateNumbering(wordDocument);
                if (numbering != null)
                {
                    part = wordDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                    part.Numbering = numbering;
                    //part.Numbering.Save();
                }
            }
            return wordDocument;
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>The transformed <see cref="OpenXmlElement" />.</returns>
        protected virtual object TransformNumbering(OpenXmlElement element, WordprocessingDocument wordDocument)
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
        /// <param name="wordDocument">The <see cref="WordprocessingDocument" /> to be transformed.</param>
        /// <returns>A new instance of <see cref="Numbering" /> or null.</returns>
        protected virtual Numbering CreateNumbering(WordprocessingDocument wordDocument)
        {
            return null;
        }

        #endregion Numbering

        #region Headers

        protected WordprocessingDocument TransformHeaders(WordprocessingDocument wordDocument)
        {
            var parts = wordDocument.MainDocumentPart.HeaderParts.ToList();
            foreach (var part in parts)
            {
                var header = (Header) TransformHeader(part.Header, wordDocument);
                if (header != null)
                {
                    part.Header = header;
                    //part.Header.Save();
                }
                else
                {
                    wordDocument.MainDocumentPart.DeletePart(part);
                }
            }
            return wordDocument;
        }

        protected virtual object TransformHeader(OpenXmlElement element, WordprocessingDocument wordDocument)
        {
            return element.CloneNode(true);
        }

        #endregion

        #region Footers

        protected WordprocessingDocument TransformFooters(WordprocessingDocument wordDocument)
        {
            var parts = wordDocument.MainDocumentPart.FooterParts.ToList();
            foreach (var part in parts)
            {
                var footer = (Footer) TransformFooter(part.Footer, wordDocument);
                if (footer != null)
                {
                    part.Footer = footer;
                    //part.Footer.Save();
                }
                else
                {
                    wordDocument.MainDocumentPart.DeletePart(part);
                }
            }
            return wordDocument;
        }

        protected virtual object TransformFooter(OpenXmlElement element, WordprocessingDocument wordDocument)
        {
            return element.CloneNode(true);
        }

        #endregion

        #region ICloneable Methods

        /// <summary>
        /// Creates a deep copy of the transformation.
        /// </summary>
        /// <returns>The clone.</returns>
        public override object Clone()
        {
            var transformation = (WordprocessingDocumentTransformation)base.Clone();
            if (Template != null)
                transformation.Template = (WordprocessingDocument)Template.Clone();

            return transformation;
        }

        #endregion
    }
}