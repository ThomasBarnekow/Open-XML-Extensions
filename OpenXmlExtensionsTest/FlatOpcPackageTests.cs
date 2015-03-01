/*
 * FlatOpcPackageTests.cs - Test driver for FlatOpcPackage
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
using System.IO.Packaging;
using System.IO.Packaging.FlatOpc;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using NUnit.Framework;

namespace OpenXmlExtensionsTest
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// This class tests the FlatOpcPackage functionality. It will be replaced
    /// by NUnit-based tests sooner or later.
    /// </summary>
    [TestFixture]
    public class FlatOpcPackageTests
    {
        static readonly string documentPath = "Document.docx";
        static readonly string spreadsheetPath = "Spreadsheet.xlsx";
        static readonly string presentationPath = "Presentation.pptx";

        static readonly string xmlDocumentPath = "Document.xml";
        // static readonly string xmlSpreadsheetPath = "Spreadsheet.xml";
        // static readonly string xmlPresentationPath = "Presentation.xml";

        static readonly string xmlCreatedDocumentPath = "FlatOpcPackage\\CreatedDocument.xml";

        static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        [TestFixtureSetUp]
        public void SetUp()
        {
            Directory.CreateDirectory("FlatOpcPackage");

            TestTools.RemoveFiles("FlatOpcPackage", "*.docx");
            TestTools.RemoveFiles("FlatOpcPackage", "*.xlsx");
            TestTools.RemoveFiles("FlatOpcPackage", "*.pptx");
            TestTools.RemoveFiles("FlatOpcPackage", "*.xml");

            File.Copy(@"..\..\" + documentPath, documentPath, true);
            File.Copy(@"..\..\" + spreadsheetPath, spreadsheetPath, true);
            File.Copy(@"..\..\" + presentationPath, presentationPath, true);

            File.Copy(@"..\..\" + xmlDocumentPath, xmlDocumentPath, true);

            TestTools.PrepareWordprocessingDocument(documentPath);
            TestTools.PrepareSpreadsheetDocument(spreadsheetPath);
            TestTools.PreparePresentationDocument(presentationPath);
        }

        private void PrintParts(Package package)
        {
            Console.WriteLine("\n" + package.GetType());
            foreach (var packageRel in package.GetRelationships().Where(r => r.TargetMode == TargetMode.Internal))
            {
                Uri partUri = PackUriHelper.CreatePartUri(packageRel.TargetUri);
                PrintPart(package.GetPart(partUri), 2);
            }
        }

        private void PrintPart(PackagePart part, int indent)
        {
            Console.WriteLine(new string(' ', indent) + part.Uri + ": " + part.ContentType);
            if (part.ContentType != "application/vnd.openxmlformats-package.relationships+xml")
            {
                foreach (var rel in part.GetRelationships())
                {
                    if (rel.TargetMode == TargetMode.Internal)
                    {
                        Uri targetPartUri = PackUriHelper.ResolvePartUri(part.Uri, rel.TargetUri);
                        PrintPart(rel.Package.GetPart(targetPartUri), indent + 2);
                    }
                    else
                    {
                        Console.WriteLine(new string(' ', indent + 2) + rel.TargetUri + " (" + rel.RelationshipType + ")");
                    }
                }
            }
        }

        [Test]
        public void ShowParts()
        {
            // This is not really an NUnit test because it doesn't assert anything.
            // However, it shows the package contents using relationships. This way,
            // we can see that the FlatOpcPackage has all the relationships although
            // we did not pay specific attention to them.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, false))
            {
                PrintParts(doc.Package);
            }

            using (FlatOpcPackage package = FlatOpcPackage.Open(xmlDocumentPath))
            {
                PrintParts(package);
            }
        }

        [Test]
        public void CreateFlatOpcPackageTest()
        {
            string text = "Hello World!";

            // Create a new package.
            FlatOpcPackage package = FlatOpcPackage.Open(xmlCreatedDocumentPath, FileMode.Create);
            using (WordprocessingDocument doc = WordprocessingDocument.Create(package,
                WordprocessingDocumentType.Document))
            {
                MainDocumentPart part = doc.AddMainDocumentPart();
                part.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text(text)))));

            }

            // Open that package again and check result.
            FlatOpcPackage testPackage = FlatOpcPackage.Open(xmlCreatedDocumentPath, FileMode.Open);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPackage))
            {
                Document document = doc.MainDocumentPart.Document;
                string testText = document.Descendants<Text>().First().Text;
                Assert.That(testText, Is.EqualTo(text));
            }
        }

        [Test]
        public void CloneWordprocessingDocumentTest()
        {
            // Let's do something bigger, i.e., clone an OPC package-based Word document,
            // to a FlatOpcPackage-based Open XML document.
            FlatOpcPackage package = FlatOpcPackage.Open(xmlCreatedDocumentPath, FileMode.Create);
            using (WordprocessingDocument original = WordprocessingDocument.Open(documentPath, false))
            using (WordprocessingDocument clone = (WordprocessingDocument)original.Clone(package))
            {
                TestTools.AssertThatPackagesAreEqual(original, clone);
            }
        }

        [Test]
        public void OpenFromXDocumentTest()
        {
            string text = "Inserted before first paragraph";

            // Let's now open the clone from an XDocument and look at the contents.
            XDocument cloneDoc = XDocument.Load(xmlCreatedDocumentPath);
            FlatOpcPackage package = FlatOpcPackage.Open(cloneDoc);
            using (WordprocessingDocument clone = WordprocessingDocument.Open(package))
            {
                Document document = clone.MainDocumentPart.Document;
                Paragraph p = document.Body.Elements<Paragraph>().First();

                p.InsertBeforeSelf(
                    new Paragraph(
                        new Run(
                            new Text(text))));
            }

            // Let's access the package's Document again to see whether that works.
            XDocument testDoc = package.Document;
            XElement testElement = testDoc.Descendants(w + "t").First();
            
            Assert.That(testElement.Value, Is.EqualTo(text));
        }
    }
}
