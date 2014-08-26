/*
 * FlatOpcTransform.cs - Transforming between OPC and Flat OPC formats
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
 * Developer: Thomas Barnekow, with input from Eric White's MSDN blog posts
 * Email: thomas<at/>barnekow<dot/>info
 * 
 * Version: 1.0.01
 */

using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transforms
{
    /// <summary>
    /// Provides a number of utility methods for working with OPC (Open Packaging
    /// Convention) formats.
    /// </summary>
    public static class FlatOpcTransform
    {
        #region OPC to Flat OPC conversion

        /// <summary>
        /// Gets the <see cref="PackagePart"/>'s contents as an <see cref="XElement"/>.
        /// </summary>
        /// <param name="part">The package part</param>
        /// <returns>The corresponding <see cref="XElement"/></returns>
        private static XElement GetContentsAsXml(PackagePart part)
        {
            // Author: Eric White. 
            // Source code taken from MSDN blog post.
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

            if (part.ContentType.EndsWith("xml"))
            {
                using (Stream stream = part.GetStream())
                using (StreamReader streamReader = new StreamReader(stream))
                using (XmlReader xmlReader = XmlReader.Create(streamReader))
                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XElement(pkg + "xmlData", XElement.Load(xmlReader)));
            }
            else
            {
                using (Stream stream = part.GetStream())
                using (BinaryReader binaryReader = new BinaryReader(stream))
                {
                    int len = (int)binaryReader.BaseStream.Length;
                    byte[] byteArray = binaryReader.ReadBytes(len);

                    // The following expression creates the base64String, then chunks
                    // it to lines of 76 characters long
                    string base64String = System.Convert.ToBase64String(byteArray)
                        .Select((c, i) => new { Character = c, Chunk = i / 76 })
                        .GroupBy(c => c.Chunk)
                        .Aggregate(
                            new StringBuilder(),
                            (s, i) => 
                                s.Append(
                                    i.Aggregate(
                                        new StringBuilder(),
                                        (seed, it) => seed.Append(it.Character),
                                        sb => sb.ToString())).Append(Environment.NewLine),
                            s => s.ToString());

                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XAttribute(pkg + "compression", "store"),
                        new XElement(pkg + "binaryData", base64String)
                    );
                }
            }
        }

        /// <summary>
        /// Processing instructions for Word document.
        /// </summary>
        private static readonly XProcessingInstruction WordDocument = 
            new XProcessingInstruction("mso-application", "progid=\"Word.Document\"");

        /// <summary>
        /// Processing instructions for PowerPoint document.
        /// </summary>
        private static readonly XProcessingInstruction PowerPointShow = 
            new XProcessingInstruction("mso-application", "progid=\"PowerPoint.Show\"");

        /// <summary>
        /// Returns the <see cref="XProcessingInstruction"/> corresponding to the 
        /// file name's extension (e.g., ".docx"). Currently, only ".docx" and 
        /// ".pptx" are supported. 
        /// </summary>
        /// <param name="fileName">The filename</param>
        /// <returns>The processing instruction element</returns>
        private static XProcessingInstruction GetProcessingInstruction(string fileName)
        {
            // Author: Eric White (with refactoring by Thomas Barnekow)
            // Source code taken from MSDN blog post.

            // Check extension and return processing instructions for Word documents
            if (fileName.ToLower().EndsWith(".docx"))
                return WordDocument;
            
            // Check extension and return processing instructions for PowerPoint presentation
            if (fileName.ToLower().EndsWith(".pptx"))
                return PowerPointShow;
            
            // Neither Word nor PowerPoint document
            return null;
        }

        /// <summary>
        /// Converts a <see cref="Package"/> in OPC format to an <see cref="XDocument"/> 
        /// in Flat OPC format.
        /// </summary>
        /// <param name="package">The package in OPC format</param>
        /// <param name="instruction">The processing instructions</param>
        /// <returns>The document in Flat OPC format</returns>
        public static XDocument ToFlatOpc(Package package, XProcessingInstruction instruction)
        {
            // Author: Eric White (with refactoring by Thomas Barnekow)
            // Source code taken from MSDN blog post.
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            
            XDocument doc =
                new XDocument(
                    new XDeclaration("1.0", "UTF-8", "yes"),
                    instruction,
                    new XElement(
                        pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        package.GetParts().Select(part => GetContentsAsXml(part))));

            return doc;
        }

        /// <summary>
        /// Converts <see cref="OpenXmlPackage"/> in OPC format to Flat OPC 
        /// format after saving the data in the DOM tree back to the parts 
        /// and flushing the package.
        /// </summary>
        /// <param name="openXmlPackage">The package in OPC format</param>
        /// <param name="instruction">The processing instructions</param>
        /// <returns>The package in Flat OPC format</returns>
        public static XDocument ToFlatOpc(OpenXmlPackage openXmlPackage, 
            XProcessingInstruction instruction)
        {
            if (openXmlPackage == null)
                throw new ArgumentNullException("openXmlPackage");

            // Save root elements of all parts contained in document to their
            // respective package parts. 
            // foreach (OpenXmlPart part in openXmlPackage.GetAllParts())
            //     part.RootElement.Save();
            openXmlPackage.Save();

            // Save all parts to package.
            // openXmlPackage.Package.Flush();

            // Convert package.
            return FlatOpcTransform.ToFlatOpc(openXmlPackage.Package, instruction);
        }

        /// <summary>
        /// Converts <see cref="WordprocessingDocument"/> in OPC format to Flat 
        /// OPC format after saving the data in the DOM tree back to the parts 
        /// and flushing the package.
        /// </summary>
        /// <param name="document">The document in OPC format</param>
        /// <returns>The document in Flat OPC format</returns>
        public static XDocument ToFlatOpc(WordprocessingDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return FlatOpcTransform.ToFlatOpc(document, WordDocument);
        }

        /// <summary>
        /// Converts <see cref="PresentationDocument"/> in OPC format to Flat 
        /// OPC format after saving the data in the DOM tree back to the parts 
        /// and flushing the package.
        /// </summary>
        /// <param name="document">The document in OPC format</param>
        /// <returns>The document in Flat OPC format</returns>
        public static XDocument ToFlatOpc(PresentationDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return FlatOpcTransform.ToFlatOpc(document, PowerPointShow);
        }

        /// <summary>
        /// Converts <see cref="SpreadsheetDocument"/> in OPC format to Flat 
        /// OPC format after saving the data in the DOM tree back to the parts 
        /// and flushing the package.
        /// </summary>
        /// <param name="document">The document in OPC format</param>
        /// <returns>The document in Flat OPC format</returns>
        public static XDocument ToFlatOpc(SpreadsheetDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return FlatOpcTransform.ToFlatOpc(document, null);
        }

        /// <summary>
        /// Converts the given (Word) file to Flat OPC format.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static XDocument ToFlatOpc(string path)
        {
            using (Package package = Package.Open(path))
                return ToFlatOpc(package, GetProcessingInstruction(path));
        }

        #endregion OPC to Flat OPC conversion

        #region Flat OPC to OPC conversion

        /// <summary>
        /// Converts an <see cref="XDocument"/> in Flat OPC format to a <see cref="Package"/> 
        /// in OPC format.
        /// </summary>
        /// <param name="doc">The document in Flat OPC format</param>
        /// <param name="package">The package in OPC format</param>
        public static void ToOpc(XDocument doc, Package package)
        {
            // Author: Eric White. 
            // Source code taken from MSDN blog post.
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

            // Add all parts (but not relationships)
            foreach (var xmlPart in doc.Root
                .Elements()
                .Where(p =>
                    (string)p.Attribute(pkg + "contentType") !=
                    "application/vnd.openxmlformats-package.relationships+xml"))
            {
                string name = (string)xmlPart.Attribute(pkg + "name");
                string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                if (contentType.EndsWith("xml"))
                {
                    Uri uri = new Uri(name, UriKind.Relative);
                    PackagePart part = package.CreatePart(uri, contentType, CompressionOption.SuperFast);
                    using (Stream stream = part.GetStream(FileMode.Create))
                    using (XmlWriter xmlWriter = XmlWriter.Create(stream))
                        xmlPart.Element(pkg + "xmlData")
                            .Elements()
                            .First()
                            .WriteTo(xmlWriter);
                }
                else
                {
                    Uri uri = new Uri(name, UriKind.Relative);
                    PackagePart part = package.CreatePart(uri, contentType, CompressionOption.SuperFast);
                    using (Stream stream = part.GetStream(FileMode.Create))
                    using (BinaryWriter binaryWriter = new BinaryWriter(stream))
                    {
                        string base64StringInChunks = (string)xmlPart.Element(pkg + "binaryData");
                        char[] base64CharArray = base64StringInChunks
                            .Where(c => c != '\r' && c != '\n').ToArray();
                        byte[] byteArray =
                            System.Convert.FromBase64CharArray(
                                base64CharArray, 0, base64CharArray.Length);
                        binaryWriter.Write(byteArray);
                    }
                }
            }

            foreach (var xmlPart in doc.Root.Elements())
            {
                string name = (string)xmlPart.Attribute(pkg + "name");
                string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                if (contentType == "application/vnd.openxmlformats-package.relationships+xml")
                {
                    if (name == "/_rels/.rels")
                    {
                        // Add the package level relationships
                        foreach (XElement xmlRel in xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode = (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External")
                                package.CreateRelationship(
                                    new Uri(target, UriKind.Absolute),
                                    TargetMode.External, type, id);
                            else
                                package.CreateRelationship(
                                    new Uri(target, UriKind.Relative),
                                    TargetMode.Internal, type, id);
                        }
                    }
                    else
                    {
                        // Add part level relationships
                        string directory = name.Substring(0, name.IndexOf("/_rels"));
                        string relsFilename = name.Substring(name.LastIndexOf('/'));
                        string filename = relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                        PackagePart fromPart = package.GetPart(new Uri(directory + filename, UriKind.Relative));
                        foreach (XElement xmlRel in xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode = (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External")
                                fromPart.CreateRelationship(
                                    new Uri(target, UriKind.Absolute),
                                    TargetMode.External, type, id);
                            else
                                fromPart.CreateRelationship(
                                    new Uri(target, UriKind.Relative),
                                    TargetMode.Internal, type, id);
                        }
                    }
                }
            }

            // Save contents of all parts and relationships contained in package
            package.Flush();
        }

        /// <summary>
        /// Converts an <see cref="XDocument"/> in Flat OPC format to a file in OPC format.
        /// </summary>
        /// <param name="doc">The document in Flat OPC format</param>
        /// <param name="path">The path of the file in OPC format</param>
        public static void ToOpc(XDocument doc, string path)
        {
            using (Package package = Package.Open(path, FileMode.Create))
                ToOpc(doc, package);
        }

        #endregion Flat OPC to OPC conversion
    }
}
