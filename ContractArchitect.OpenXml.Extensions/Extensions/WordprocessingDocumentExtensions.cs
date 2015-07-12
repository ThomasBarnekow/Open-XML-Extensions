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
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using ContractArchitect.OpenXml.Transformation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for <see cref="WordprocessingDocument" /> class.
    /// </summary>
    [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
    public static class WordprocessingDocumentExtensions
    {
        /// <summary>
        /// Binds content controls to a custom XML part created or updated from the given XML document.
        /// </summary>
        /// <param name="document">The WordprocessingDocument.</param>
        /// <param name="rootElement">The custom XML part's root element.</param>
        public static void BindContentControls(this WordprocessingDocument document, XElement rootElement)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (rootElement == null)
                throw new ArgumentNullException("rootElement");

            // Get or create custom XML part. This assumes that we only have a single custom
            // XML part for any given namespace.
            var destPart = document.GetCustomXmlPart(rootElement.Name.Namespace);
            if (destPart == null)
                destPart = document.CreateCustomXmlPart(rootElement);
            else
                destPart.SetRootElement(rootElement);

            // Bind the content controls to the destination part's XML document.
            document.BindContentControls(destPart);
        }

        /// <summary>
        /// Binds content controls to a custom XML part.
        /// </summary>
        /// <param name="document">The WordprocessingDocument.</param>
        /// <param name="part">The custom XML part.</param>
        public static void BindContentControls(this WordprocessingDocument document, CustomXmlPart part)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (part == null)
                throw new ArgumentNullException("part");

            var customXmlRootElement = part.GetRootElement();
            var storeItemId = part.CustomXmlPropertiesPart.DataStoreItem.ItemId.Value;

            // Bind w:sdt elements contained in main document part.
            var partRootElement = document.MainDocumentPart.RootElement;
            BindContentControls(partRootElement, customXmlRootElement, storeItemId);
            partRootElement.Save();

            // Bind w:sdt elements contained in header parts.
            foreach (var headerRootElement in document.MainDocumentPart
                .HeaderParts.Select(p => p.RootElement))
            {
                BindContentControls(headerRootElement, customXmlRootElement, storeItemId);
                headerRootElement.Save();
            }

            // Bind w:sdt elements contained in footer parts.
            foreach (var footerRootElement in document.MainDocumentPart
                .FooterParts.Select(p => p.RootElement))
            {
                BindContentControls(footerRootElement, customXmlRootElement, storeItemId);
                footerRootElement.Save();
            }
        }

        /// <summary>
        /// Bind the content controls (w:sdt elements) contained in the content part's XML document to the
        /// custom XML part identified by the given storeItemId.
        /// </summary>
        /// <param name="contentRootElement">The content part's <see cref="OpenXmlPartRootElement" />.</param>
        /// <param name="customXmlRootElement">The custom XML part's root <see cref="XElement" />.</param>
        /// <param name="storeItemId">The w:storeItemId to be used for data binding.</param>
        public static void BindContentControls(OpenXmlPartRootElement contentRootElement,
            XElement customXmlRootElement, string storeItemId)
        {
            if (contentRootElement == null)
                throw new ArgumentNullException("contentRootElement");
            if (customXmlRootElement == null)
                throw new ArgumentNullException("customXmlRootElement");
            if (storeItemId == null)
                throw new ArgumentNullException("storeItemId");

            // Get all w:sdt elements with matching tags.
            var tags = customXmlRootElement.Descendants()
                .Where(e => !e.HasElements)
                .Select(e => e.Name.LocalName);
            var sdts = contentRootElement.Descendants<SdtElement>()
                .Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null &&
                              tags.Contains(sdt.SdtProperties.GetFirstChild<Tag>().Val.Value));

            foreach (var sdt in sdts)
            {
                // The tag value is supposed to point to a descendant element of the custom XML
                // part's root element.
                var childElementName = sdt.SdtProperties.GetFirstChild<Tag>().Val.Value;
                var leafElement = customXmlRootElement.Descendants()
                    .First(e => e.Name.LocalName == childElementName);

                // Define the list of path elements, using one of the following two options:
                // 1. The following statement is used as the basis for building the full path
                // expression (the same as built by Microsoft Word).
                var pathElements = leafElement.AncestorsAndSelf().Reverse().ToList();

                // 2. The following statement is used as the basis for building the short xPath
                // expression "//ns0:leafElement[1]".
                // List<XElement> pathElements = new List<XElement>() { leafElement };

                // Build list of namespace names for building the prefix mapping later on.
                var nsList = pathElements
                    .Where(e => e.Name.Namespace != XNamespace.None)
                    .Aggregate(new HashSet<string>(), (set, e) => set.Append(e.Name.NamespaceName))
                    .ToList();

                // Build mapping from local names to namespace indices.
                var nsDict = pathElements
                    .ToDictionary(e => e.Name.LocalName, e => nsList.IndexOf(e.Name.NamespaceName));

                // Build prefix mappings.
                var prefixMappings = nsList.Select((ns, index) => new {ns, index})
                    .Aggregate(new StringBuilder(), (sb, t) =>
                        sb.Append("xmlns:ns").Append(t.index).Append("='").Append(t.ns).Append("' "))
                    .ToString().Trim();

                // Build xPath, assuming we will always take the first element and using one
                // of the following two options (see above):
                // 1. The following statement defines the prefix for building a full path
                // expression "/ns0:path[1]/ns0:to[1]/ns0:leafElement[1]".
                Func<string, string> prefix = localName =>
                    nsDict[localName] >= 0 ? "/ns" + nsDict[localName] + ":" : "/";

                // 2. The following statement defines the prefix for building the short path
                // expression "//ns0:leafElement[1]".
                // Func<string, string> prefix = localName =>
                //     nsDict[localName] >= 0 ? "//ns" + nsDict[localName] + ":" : "//";

                var xPath = pathElements
                    .Select(e => prefix(e.Name.LocalName) + e.Name.LocalName + "[1]")
                    .Aggregate(new StringBuilder(), (sb, pc) => sb.Append(pc)).ToString();

                // Create and configure new data binding.
                var dataBinding = new DataBinding();
                if (!String.IsNullOrEmpty(prefixMappings))
                    dataBinding.PrefixMappings = prefixMappings;
                dataBinding.XPath = xPath;
                dataBinding.StoreItemId = storeItemId;

                // Add or replace data binding.
                var currentDataBinding = sdt.SdtProperties.GetFirstChild<DataBinding>();
                if (currentDataBinding != null)
                    sdt.SdtProperties.ReplaceChild(dataBinding, currentDataBinding);
                else
                    sdt.SdtProperties.Append(dataBinding);
            }
        }

        /// <summary>
        /// Creates a <see cref="CustomXmlPart" /> with the given root <see cref="XElement" />.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="rootElement">The root element.</param>
        /// <returns>The newly created custom XML part.</returns>
        public static CustomXmlPart CreateCustomXmlPart(this WordprocessingDocument document, XElement rootElement)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (rootElement == null)
                throw new ArgumentNullException("rootElement");

            // Create a ds:dataStoreItem associated with the custom XML part's root element.
            var dataStoreItem = new DataStoreItem
            {
                ItemId = "{" + Guid.NewGuid().ToString().ToUpper() + "}",
                SchemaReferences = new SchemaReferences()
            };
            if (rootElement.Name.Namespace != XNamespace.None)
                dataStoreItem.SchemaReferences.Append(new SchemaReference {Uri = rootElement.Name.NamespaceName});

            // Create the custom XML part.
            var customXmlPart = document.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPart.SetRootElement(rootElement);

            // Create the custom XML properties part.
            var propertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propertiesPart.DataStoreItem = dataStoreItem;
            propertiesPart.DataStoreItem.Save();

            // Commented out after running into issues with ZipStreamManager.
            // document.Package.Flush();
            return customXmlPart;
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
            var style = document.GetParagraphStyle(styleId);
            if (style != null)
                throw new ArgumentException("Style '" + styleId + "' already exists!", styleId);

            // Create a new paragraph style element and specify key attributes.
            style = new Style {Type = StyleValues.Paragraph, CustomStyle = true, StyleId = styleId};

            // Add key child elements
            style.Produce<StyleName>().Val = styleName;
            style.Produce<BasedOn>().Val = basedOn;
            style.Produce<NextParagraphStyle>().Val = nextStyle;

            // Add the style to the styles part
            return document.ProduceStylesElement().AppendChild(style);
        }

        /// <summary>
        /// Gets the character <see cref="Style" /> with the given id.
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

            var styles = document.ProduceStylesElement();
            return styles.Elements<Style>().FirstOrDefault<Style>(
                style => style.StyleId == styleId &&
                         style.Type == StyleValues.Character);
        }

        /// <summary>
        /// Returns the <see cref="CustomXmlPart" /> having a root element with the given <see cref="XNamespace" />
        /// or null if there is no such <see cref="CustomXmlPart" />.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="ns">The namespace</param>
        /// <returns>The corresponding part or null</returns>
        public static CustomXmlPart GetCustomXmlPart(this WordprocessingDocument document, XNamespace ns)
        {
            if (document != null && document.MainDocumentPart != null)
                return document.MainDocumentPart
                    .CustomXmlParts
                    .LastOrDefault(p => p.GetRootNamespace() == ns);

            return null;
        }

        /// <summary>
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultFontSize(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var rPr = document.GetRunPropertiesBaseStyle();
            return rPr.FontSize != null ? int.Parse(rPr.FontSize.Val) : 20;
        }

        /// <summary>
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultSpaceAfter(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var pPr = document.GetParagraphPropertiesBaseStyle();
            if (pPr.SpacingBetweenLines == null)
                return 0;
            if (pPr.SpacingBetweenLines.After == null)
                return 0;

            return int.Parse(pPr.SpacingBetweenLines.After);
        }

        /// <summary>
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetDefaultSpaceBefore(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var pPr = document.GetParagraphPropertiesBaseStyle();
            if (pPr.SpacingBetweenLines == null)
                return 0;
            if (pPr.SpacingBetweenLines.Before == null)
                return 0;

            return int.Parse(pPr.SpacingBetweenLines.Before);
        }

        /// <summary>
        /// Gets the last section's <see cref="PageMargin" /> element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="PageMargin" /> element</returns>
        public static PageMargin GetPageMargin(this WordprocessingDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return document.GetSectionProperties().GetFirstChild<PageMargin>();
        }

        /// <summary>
        /// Gets the last section's <see cref="PageSize" /> element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="PageSize" /> element</returns>
        public static PageSize GetPageSize(this WordprocessingDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return document.GetSectionProperties().GetFirstChild<PageSize>();
        }

        /// <summary>
        /// Gets the <see cref="ParagraphPropertiesBaseStyle" /> ancestor of the document's styles element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="ParagraphPropertiesBaseStyle" /></returns>
        public static ParagraphPropertiesBaseStyle GetParagraphPropertiesBaseStyle(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var docDefaults = document.ProduceStylesElement().DocDefaults;
            return docDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle;
        }

        /// <summary>
        /// Gets the paragraph <see cref="Style" /> with the given id.
        /// </summary>
        /// <param name="document">The document</param>
        /// <param name="styleId">The style's id</param>
        /// <returns>The corresponding style</returns>
        public static Style GetParagraphStyle(this WordprocessingDocument document, string styleId)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (styleId == null)
                throw new ArgumentNullException("styleId");

            var styles = document.ProduceStylesElement();
            return styles.Elements<Style>().FirstOrDefault<Style>(
                style => style.StyleId == styleId &&
                         style.Type == StyleValues.Paragraph);
        }

        /// <summary>
        /// Gets the <see cref="RunPropertiesBaseStyle" /> ancestor of the document's styles element.
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The <see cref="RunPropertiesBaseStyle" /></returns>
        public static RunPropertiesBaseStyle GetRunPropertiesBaseStyle(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var docDefaults = document.ProduceStylesElement().DocDefaults;
            return docDefaults.RunPropertiesDefault.RunPropertiesBaseStyle;
        }

        /// <summary>
        /// Gets the last section's properties
        /// </summary>
        /// <param name="document">The document</param>
        /// <returns>The last section's <see cref="SectionProperties" /> element</returns>
        public static SectionProperties GetSectionProperties(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            // The body's SectionProperties element represents the last section's properties.
            return document.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
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

            var part = document.MainDocumentPart.NumberingDefinitionsPart ??
                       document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

            if (part.Numbering != null)
                return part.Numbering;

            part.Numbering = new Numbering();
            part.Numbering.Save();
            return part.Numbering;
        }

        public static Settings ProduceSettingsElement(this WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            var part = document.MainDocumentPart.DocumentSettingsPart ??
                       document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();

            if (part.Settings != null)
                return part.Settings;

            part.Settings = new Settings();
            part.Settings.Save();
            return part.Settings;
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

            var part = document.MainDocumentPart.StyleDefinitionsPart ??
                       document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            if (part.Styles != null)
                return part.Styles;

            part.Styles = new Styles();
            part.Styles.Save();
            return part.Styles;
        }
    }
}
