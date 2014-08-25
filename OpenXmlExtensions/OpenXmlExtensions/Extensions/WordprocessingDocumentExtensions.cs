/*
 * WordprocessingDocumentExtensions.cs - Extensions for WordprocessingDocument
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
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for <see cref="WordprocessingDocument"/> class.
    /// </summary>
    public static class WordprocessingDocumentExtensions
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

            // Copy all document parts (AddPart will copy the parts and their 
            // children in a recursive fashion).
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

        /// <summary>
        /// Returns the <see cref="CustomXmlPart"/> having a root element with the given <see cref="XNamespace"/> 
        /// or null if there is no such <see cref="CustomXmlPart"/>.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="ns">The namespace</param>
        /// <returns>The corresponding part or null</returns>
        public static CustomXmlPart GetCustomXmlPart(this WordprocessingDocument document, XNamespace ns)
        {
            if (document != null && document.MainDocumentPart != null)
                return document.MainDocumentPart
                    .GetPartsOfType<CustomXmlPart>()
                    .SingleOrDefault<CustomXmlPart>(p => p.GetRootNamespace() == ns);
            else
                return null;
        }

        /// <summary>
        /// Creates a <see cref="CustomXmlPart"/> with the given root <see cref="XElement"/>.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="partRoot">The root element</param>
        /// <returns>The newly created custom XML part</returns>
        public static CustomXmlPart CreateCustomXmlPart(this WordprocessingDocument document, XElement partRoot)
        {
            if (partRoot == null)
                return null;

            // Add custom XML part
            CustomXmlPart part = document.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            part.SetRootElement(partRoot);

            // Create contents of XML properties part
            XNamespace ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
            XElement propertyPartRoot = new XElement(ds + "datastoreItem",
                new XAttribute(ds + "itemID", "{" + Guid.NewGuid().ToString().ToUpper() + "}"),
                new XAttribute(XNamespace.Xmlns + "ds", ds.NamespaceName),
                new XElement(ds + "schemaRefs"));

            // Add custom XML properties part
            CustomXmlPropertiesPart propertyPart = part.AddNewPart<CustomXmlPropertiesPart>();
            propertyPart.SetRootElement(propertyPartRoot);

            // Done
            return part;
        }

        /// <summary>
        /// Gets or creates the root element of the document's style definitions part.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The root element of the document's style definitions part</returns>
        public static Styles ProduceStylesElement(this WordprocessingDocument document)
        {
            // Access the styles part.
            StyleDefinitionsPart part = document.MainDocumentPart.StyleDefinitionsPart;
            if (part == null)
                part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            // Access the root element of the styles part.
            if (part.Styles == null)
            {
                // Create and save the root element
                part.Styles = new Styles();
                part.Styles.Save();
                document.Package.Flush();
            }

            // Done.
            return part.Styles;
        }

        /// <summary>
        /// Gets or crates the root element of the document's numbering definitions part.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The root element of the document's numbering definitions part</returns>
        public static Numbering ProduceNumberingElement(this WordprocessingDocument document)
        {
            // Access the numbering part.
            NumberingDefinitionsPart part = document.MainDocumentPart.NumberingDefinitionsPart;
            if (part == null)
                part = document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

            // Access the root element of the numbering part.
            if (part.Numbering == null)
            {
                // Create and save the root element
                part.Numbering = new Numbering();
                part.Numbering.Save();
                document.Package.Flush();
            }

            // Done.
            return part.Numbering;
        }
 
        /// <summary>
        /// Gets the paragraph <see cref="Style"/> with the given id.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="styleId">The style's id</param>
        /// <returns>The corresponding style</returns>
        public static Style GetParagraphStyle(this WordprocessingDocument document, 
            string styleId)
        {
            Styles styles = document.ProduceStylesElement();
            return styles.Elements<Style>().FirstOrDefault<Style>(
                style => style.StyleId == styleId &&
                         style.Type == StyleValues.Paragraph);
        }

        /// <summary>
        /// Gets the character <see cref="Style"/> with the given id.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="styleId">The style's id</param>
        /// <returns>The corresponding style</returns>
        public static Style GetCharacterStyle(this WordprocessingDocument document, 
            string styleId)
        {
            Styles styles = document.ProduceStylesElement();
            return styles.Elements<Style>().FirstOrDefault<Style>(
                style => style.StyleId == styleId &&
                         style.Type == StyleValues.Character);
        }

        /// <summary>
        /// Creates a new paragraph style with the specified style ID, primary 
        /// style name, and aliases and add it to the specified style definitions
        /// part. Saves the data in the DOM tree back to the part.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="styleId">The style's unique ID</param>
        /// <param name="styleName">The style's name</param>
        /// <param name="basedOn">The base style</param>
        /// <param name="nextStyle">The next paragraph's style</param>
        /// <returns>The newly created style</returns>
        public static Style CreateParagraphStyle(this WordprocessingDocument document, 
            string styleId, string styleName, string basedOn, string nextStyle)
        {
            // Check parameters
            if (document == null)
                throw new ArgumentNullException("document");
            if (styleId == null)
                throw new ArgumentNullException("styleId");
            if (styleName == null)
                throw new ArgumentNullException("styleName");
            if (basedOn == null)
                throw new ArgumentNullException("basedOn");
            if (nextStyle == null)
                throw new ArgumentNullException("nextStyle");

            // Check whether the style already exists.
            Style style = document.GetParagraphStyle(styleId);
            if (style != null)
                throw new ArgumentException("Style '" + styleId + "' already exists!", styleId);

            // Create a new paragraph style element and specify key attributes.
            style = new Style() { Type = StyleValues.Paragraph, CustomStyle = true, StyleId = styleId };

            // Add key child elements
            style.Produce<StyleName>().Val = styleName;
            style.Produce<BasedOn>().Val = basedOn;
            style.Produce<NextParagraphStyle>().Val = nextStyle;

            // Add the style to the styles part
            return document.ProduceStylesElement().AppendChild<Style>(style);
        }

        /// <summary>
        /// Gets the last section's properties
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The last section's <see cref="SectionProperties"/> element</returns>
        public static SectionProperties GetSectionProperties(this WordprocessingDocument document)
        {
            // Check prerequisites
            if (document == null) 
                throw new ArgumentNullException("document");

            // The body's SectionProperties element represents the last section's properties
            return document.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
        }

        /// <summary>
        /// Gets the last section's <see cref="PageMargin"/> element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="PageMargin"/> element</returns>
        public static PageMargin GetPageMargin(this WordprocessingDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return document.GetSectionProperties().GetFirstChild<PageMargin>();
        }

        /// <summary>
        /// Gets the last section's <see cref="PageSize"/> element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="PageSize"/> element</returns>
        public static PageSize GetPageSize(this WordprocessingDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return document.GetSectionProperties().GetFirstChild<PageSize>();
        }

        /// <summary>
        /// Gets the <see cref="ParagraphPropertiesBaseStyle"/> ancestor of the document's styles element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="ParagraphPropertiesBaseStyle"/></returns>
        public static ParagraphPropertiesBaseStyle GetParagraphPropertiesBaseStyle(this WordprocessingDocument document)
        {
            DocDefaults docDefaults = document.ProduceStylesElement().DocDefaults;
            return docDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultSpaceBefore(this WordprocessingDocument document)
        {
            ParagraphPropertiesBaseStyle pPr = document.GetParagraphPropertiesBaseStyle();
            if (pPr.SpacingBetweenLines == null)
                return 0;
            if (pPr.SpacingBetweenLines.Before == null)
                return 0;

            return int.Parse(pPr.SpacingBetweenLines.Before);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultSpaceAfter(this WordprocessingDocument document)
        {
            ParagraphPropertiesBaseStyle pPr = document.GetParagraphPropertiesBaseStyle();
            if (pPr.SpacingBetweenLines == null)
                return 0;
            if (pPr.SpacingBetweenLines.After == null)
                return 0;

            return int.Parse(pPr.SpacingBetweenLines.After);
        }

        /// <summary>
        /// Gets the <see cref="RunPropertiesBaseStyle"/> ancestor of the document's styles element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="RunPropertiesBaseStyle"/></returns>
        public static RunPropertiesBaseStyle GetRunPropertiesBaseStyle(this WordprocessingDocument document)
        {
            DocDefaults docDefaults = document.ProduceStylesElement().DocDefaults;
            return docDefaults.RunPropertiesDefault.RunPropertiesBaseStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultFontSize(this WordprocessingDocument document)
        {
            RunPropertiesBaseStyle rPr = document.GetRunPropertiesBaseStyle();
            if (rPr.FontSize != null)
                return int.Parse(rPr.FontSize.Val);
            else
                return 20;
        }
    }
}
