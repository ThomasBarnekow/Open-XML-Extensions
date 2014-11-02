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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.CustomXmlSchemaReferences;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Transforms;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for <see cref="WordprocessingDocument"/> class.
    /// </summary>
    public static class WordprocessingDocumentExtensions
    {
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
                    .SingleOrDefault(p => p.GetRootNamespace() == ns);
            else
                return null;
        }

        /// <summary>
        /// Creates a <see cref="CustomXmlPart"/> with the given root <see cref="XElement"/>.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="root">The root element.</param>
        /// <returns>The newly created custom XML part.</returns>
        public static CustomXmlPart CreateCustomXmlPart(this WordprocessingDocument document, XElement root)
        {
            // Create a ds:dataStoreItem associated with the custom XML part's root element.
            DataStoreItem dataStoreItem = new DataStoreItem();
            dataStoreItem.ItemId = "{" + Guid.NewGuid().ToString().ToUpper() + "}";
            dataStoreItem.SchemaReferences = new SchemaReferences();
            if (root.Name.Namespace != XNamespace.None)
                dataStoreItem.SchemaReferences.Append(new SchemaReference { Uri = root.Name.NamespaceName });

            // Create the custom XML part.
            CustomXmlPart customXmlPart = document.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPart.SetRootElement(root);

            // Create the custom XML properties part.
            CustomXmlPropertiesPart propertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propertiesPart.DataStoreItem = dataStoreItem;
            propertiesPart.DataStoreItem.Save();

            document.Package.Flush();
            return customXmlPart;
        }

        /// <summary>
        /// Binds content controls to a custom XML part created or updated from the given XML document.
        /// </summary>
        /// <param name="document">The WordprocessingDocument.</param>
        /// <param name="partRootElement">The custom XML part's root element.</param>
        public static void BindContentControls(this WordprocessingDocument document, XElement partRootElement)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (partRootElement == null)
                throw new ArgumentNullException("partRootElement");

            // Get or create custom XML part. This assumes that we only have a single custom
            // XML part for any given namespace.
            CustomXmlPart destPart = document.GetCustomXmlPart(partRootElement.Name.Namespace);
            if (destPart == null)
                destPart = document.CreateCustomXmlPart(partRootElement);
            else
                destPart.SetRootElement(partRootElement);

            // Bind the content controls to the destination part's XML document.
            document.BindContentControls(destPart);
        }

        /// <summary>
        /// Binds content controls to a custom XML part.
        /// </summary>
        /// <param name="document">The WordprocessingDocument.</param>
        /// <param name="destPart">The custom XML part.</param>
        public static void BindContentControls(this WordprocessingDocument document, CustomXmlPart destPart)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (destPart == null)
                throw new ArgumentNullException("destPart");

            XElement destRoot = destPart.GetRootElement();
            string storeItemId = destPart.CustomXmlPropertiesPart.DataStoreItem.ItemId.Value;

            // Bind w:sdt elements contained in main document part.
            OpenXmlPartRootElement partRootElement = document.MainDocumentPart.RootElement;
            BindContentControls(partRootElement, destRoot, storeItemId);
            partRootElement.Save();

            // Bind w:sdt elements contained in header parts.
            foreach (OpenXmlPartRootElement headerRootElement in document.MainDocumentPart
                .HeaderParts.Select(p => p.RootElement))
            {
                BindContentControls(headerRootElement, destRoot, storeItemId);
                headerRootElement.Save();
            }

            // Bind w:sdt elements contained in footer parts.
            foreach (OpenXmlPartRootElement footerRootElement in document.MainDocumentPart
                .FooterParts.Select(p => p.RootElement))
            {
                BindContentControls(footerRootElement, destRoot, storeItemId);
                footerRootElement.Save();
            }
        }

        /// <summary>
        /// Bind the content controls (w:sdt elements) contained in the part's XML document to the
        /// custom XML part identified by the given storeItemId. 
        /// </summary>
        /// <param name="partRootElement">The Open XML part's root element.</param>
        /// <param name="storeItemId">The w:storeItemId to be used for data binding.</param>
        public static void BindContentControls(OpenXmlPartRootElement partRootElement, 
            XElement destRoot, string storeItemId)
        {
            if (partRootElement == null)
                throw new ArgumentNullException("partRootElement");
            if (storeItemId == null)
                throw new ArgumentNullException("storeItemId");
            
            // Get all w:sdt elements with matching tags.
            IEnumerable<string> tags = destRoot.Descendants().Select(e => e.Name.LocalName);
            IEnumerable<SdtElement> sdts = partRootElement.Descendants<SdtElement>()
                .Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null &&
                              tags.Contains(sdt.SdtProperties.GetFirstChild<Tag>().Val.Value));

            foreach (SdtElement sdt in sdts)
            {
                // The tag value is supposed to point to a descendant element of the custom XML
                // part's root element.
                string childElementName = sdt.SdtProperties.GetFirstChild<Tag>().Val.Value;
                XElement leafElement = destRoot.Descendants().First(e => e.Name.LocalName == childElementName);

                // Build list of namespace names for building the prefix mapping later on.
                List<XElement> pathElements = leafElement.AncestorsAndSelf().Reverse().ToList();
                List<string> nsList = pathElements
                    .Where(e => e.Name.Namespace != XNamespace.None)
                    .Aggregate(new HashSet<string>(), (set, e) => set.Append(e.Name.NamespaceName))
                    .ToList();

                // Build mapping from local names to namespace indices.
                Dictionary<string, int> nsDict = pathElements
                    .ToDictionary(e => e.Name.LocalName, e => nsList.IndexOf(e.Name.NamespaceName));

                // Build prefix mappings.
                string prefixMappings = nsList.Select((ns, index) => new { ns, index })
                    .Aggregate(new StringBuilder(), (sb, t) =>
                        sb.Append("xmlns:ns").Append(t.index).Append("='").Append(t.ns).Append("' "))
                    .ToString().Trim();

                // Build xPath, assuming we will always take the first element.
                Func<string, string> prefix = localName =>
                    nsDict[localName] >= 0 ? "/ns" + nsDict[localName] + ":" : "/";
                string xPath = pathElements
                    .Select(e => prefix(e.Name.LocalName) + e.Name.LocalName + "[1]")
                    .Aggregate(new StringBuilder(), (sb, pc) => sb.Append(pc)).ToString();

                // Create and configure new data binding.
                DataBinding dataBinding = new DataBinding();
                if (!String.IsNullOrEmpty(prefixMappings))
                    dataBinding.PrefixMappings = prefixMappings;
                dataBinding.XPath = xPath;
                dataBinding.StoreItemId = storeItemId;

                // Add or replace data binding.
                DataBinding currentDataBinding = sdt.SdtProperties.GetFirstChild<DataBinding>();
                if (currentDataBinding != null)
                    sdt.SdtProperties.ReplaceChild(dataBinding, currentDataBinding);
                else
                    sdt.SdtProperties.Append(dataBinding);
            }
        }


        /// <summary>
        /// Gets or creates the root element of the document's style definitions part.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The root element of the document's style definitions part</returns>
        public static Styles ProduceStylesElement(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

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
            if (document == null)
                throw new ArgumentNullException("document");

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
            if (document == null)
                throw new ArgumentNullException("document");
            if (styleId == null)
                throw new ArgumentNullException("styleId");

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
            if (document == null)
                throw new ArgumentNullException("document");
            if (styleId == null)
                throw new ArgumentNullException("styleId");

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
            if (document == null) 
                throw new ArgumentNullException("document");
            
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
            if (document == null) 
                throw new ArgumentNullException("document");
            
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
            if (document == null) 
                throw new ArgumentNullException("document");

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
            if (document == null) 
                throw new ArgumentNullException("document");
            
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
            if (document == null) 
                throw new ArgumentNullException("document");

            RunPropertiesBaseStyle rPr = document.GetRunPropertiesBaseStyle();
            if (rPr.FontSize != null)
                return int.Parse(rPr.FontSize.Val);
            else
                return 20;
        }
    }
}
