/*
 * FlatOpcTest.cs - Testing Flat OPC functionality in Open XML SDK
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
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

using NUnit.Framework;

namespace OpenXmlExtensionsTest
{
    using DocumentFormat.OpenXml.Wordprocessing;

    [TestFixture]
    public class FlatOpcTests
    {
        static readonly string documentPath = "Document.docx";
        static readonly string spreadsheetPath = "Spreadsheet.xlsx";
        static readonly string presentationPath = "Presentation.pptx";

        static readonly string documentClonePath = "FlatOpc\\Document Clone.docx";
        static readonly string spreadsheetClonePath = "FlatOpc\\Spreadsheet Clone.xlsx";
        static readonly string presentationClonePath = "FlatOpc\\Presentation Clone.pptx";

        [TestFixtureSetUp]
        public void SetUp()
        {
            Directory.CreateDirectory("FlatOpc");

            TestTools.RemoveFiles("FlatOpc", "*.docx");
            TestTools.RemoveFiles("FlatOpc", "*.xlsx");
            TestTools.RemoveFiles("FlatOpc", "*.pptx");

            File.Copy(@"..\..\" + documentPath, documentPath, true);
            File.Copy(@"..\..\" + spreadsheetPath, spreadsheetPath, true);
            File.Copy(@"..\..\" + presentationPath, presentationPath, true);

            TestTools.PrepareWordprocessingDocument(documentPath);
            TestTools.PrepareSpreadsheetDocument(spreadsheetPath);
            TestTools.PreparePresentationDocument(presentationPath);
        }

        [Test]
        public void TestFlatOpcWordprocessingDocument()
        {
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            {
                // Test FlatOpcDocument methods. 
                // Check ToFlatOpcDocument() and FromFlatOpcDocument(XDocument).
                XDocument flatOpcDoc = source.ToFlatOpcDocument();
                using (WordprocessingDocument dest = WordprocessingDocument.FromFlatOpcDocument(flatOpcDoc))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (WordprocessingDocument intermediate = WordprocessingDocument.FromFlatOpcDocument(flatOpcDoc, stream, false))
                using (WordprocessingDocument dest = WordprocessingDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, string, bool).
                using (WordprocessingDocument intermediate = WordprocessingDocument.FromFlatOpcDocument(flatOpcDoc, documentClonePath, false))
                using (WordprocessingDocument dest = WordprocessingDocument.Open(documentClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (WordprocessingDocument dest = WordprocessingDocument.FromFlatOpcDocument(flatOpcDoc, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Test FlatOpcString methods.
                // Check ToFlatOpcString() and FromFlatOpcString(string).
                string flatOpcString = source.ToFlatOpcString();
                using (WordprocessingDocument dest = WordprocessingDocument.FromFlatOpcString(flatOpcString))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (WordprocessingDocument intermediate = WordprocessingDocument.FromFlatOpcString(flatOpcString, stream, false))
                using (WordprocessingDocument dest = WordprocessingDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, string, bool).
                using (WordprocessingDocument intermediate = WordprocessingDocument.FromFlatOpcString(flatOpcString, documentClonePath, false))
                using (WordprocessingDocument dest = WordprocessingDocument.Open(documentClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (WordprocessingDocument dest = WordprocessingDocument.FromFlatOpcString(flatOpcString, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }

        [Test]
        public void TestFlatOpcSpreadsheetDocument()
        {
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            {
                // Test FlatOpcDocument methods. 
                // Check ToFlatOpcDocument() and FromFlatOpcDocument(XDocument).
                XDocument flatOpcDoc = source.ToFlatOpcDocument();
                using (SpreadsheetDocument dest = SpreadsheetDocument.FromFlatOpcDocument(flatOpcDoc))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (SpreadsheetDocument intermediate = SpreadsheetDocument.FromFlatOpcDocument(flatOpcDoc, stream, false))
                using (SpreadsheetDocument dest = SpreadsheetDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, string, bool).
                using (SpreadsheetDocument intermediate = SpreadsheetDocument.FromFlatOpcDocument(flatOpcDoc, spreadsheetClonePath, false))
                using (SpreadsheetDocument dest = SpreadsheetDocument.Open(spreadsheetClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (SpreadsheetDocument dest = SpreadsheetDocument.FromFlatOpcDocument(flatOpcDoc, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Test FlatOpcString methods.
                // Check ToFlatOpcString() and FromFlatOpcString(string).
                string flatOpcString = source.ToFlatOpcString();
                using (SpreadsheetDocument dest = SpreadsheetDocument.FromFlatOpcString(flatOpcString))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (SpreadsheetDocument intermediate = SpreadsheetDocument.FromFlatOpcString(flatOpcString, stream, false))
                using (SpreadsheetDocument dest = SpreadsheetDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, string, bool).
                using (SpreadsheetDocument intermediate = SpreadsheetDocument.FromFlatOpcString(flatOpcString, spreadsheetClonePath, false))
                using (SpreadsheetDocument dest = SpreadsheetDocument.Open(spreadsheetClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (SpreadsheetDocument dest = SpreadsheetDocument.FromFlatOpcString(flatOpcString, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }

        [Test]
        public void TestFlatOpcPresentationDocument()
        {
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            {
                // Test FlatOpcDocument methods. 
                // Check ToFlatOpcDocument() and FromFlatOpcDocument(XDocument).
                XDocument flatOpcDoc = source.ToFlatOpcDocument();
                using (PresentationDocument dest = PresentationDocument.FromFlatOpcDocument(flatOpcDoc))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (PresentationDocument intermediate = PresentationDocument.FromFlatOpcDocument(flatOpcDoc, stream, false))
                using (PresentationDocument dest = PresentationDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, string, bool).
                using (PresentationDocument intermediate = PresentationDocument.FromFlatOpcDocument(flatOpcDoc, presentationClonePath, false))
                using (PresentationDocument dest = PresentationDocument.Open(presentationClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcDocument(XDocument, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (PresentationDocument dest = PresentationDocument.FromFlatOpcDocument(flatOpcDoc, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Test FlatOpcString methods.
                // Check ToFlatOpcString() and FromFlatOpcString(string).
                string flatOpcString = source.ToFlatOpcString();
                using (PresentationDocument dest = PresentationDocument.FromFlatOpcString(flatOpcString))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Stream, bool).
                using (Stream stream = new MemoryStream())
                using (PresentationDocument intermediate = PresentationDocument.FromFlatOpcString(flatOpcString, stream, false))
                using (PresentationDocument dest = PresentationDocument.Open(stream, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, string, bool).
                using (PresentationDocument intermediate = PresentationDocument.FromFlatOpcString(flatOpcString, presentationClonePath, false))
                using (PresentationDocument dest = PresentationDocument.Open(presentationClonePath, false))
                    TestTools.AssertThatPackagesAreEqual(source, dest);

                // Check FromFlatOpcString(string, Package).
                using (MemoryStream stream = new MemoryStream())
                using (Package package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                using (PresentationDocument dest = PresentationDocument.FromFlatOpcString(flatOpcString, package))
                    TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }    
    }
}
