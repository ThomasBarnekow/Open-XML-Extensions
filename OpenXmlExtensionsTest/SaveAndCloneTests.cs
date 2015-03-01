/*
 * SaveAndCloneTest.cs - Testing Save and Clone functionality in Open XML SDK
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
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXmlExtensionsTest
{
    [TestFixture]
    public class SaveAndCloneTests
    {
        static readonly string documentPath = "Document.docx";
        static readonly string spreadsheetPath = "Spreadsheet.xlsx";
        static readonly string presentationPath = "Presentation.pptx";

        [TestFixtureSetUp]
        public void SetUp()
        {
            Directory.CreateDirectory("SaveAndClone");

            TestTools.RemoveFiles("SaveAndClone", "*.docx");
            TestTools.RemoveFiles("SaveAndClone", "*.xlsx");
            TestTools.RemoveFiles("SaveAndClone", "*.pptx");

            File.Copy(@"..\..\" + documentPath, documentPath, true);
            File.Copy(@"..\..\" + spreadsheetPath, spreadsheetPath, true);
            File.Copy(@"..\..\" + presentationPath, presentationPath, true);

            TestTools.PrepareWordprocessingDocument(documentPath);
            TestTools.PrepareSpreadsheetDocument(spreadsheetPath);
            TestTools.PreparePresentationDocument(presentationPath);
        }

        private void CheckWordprocessingDocument(string path, string clonePath)
        {
            using (WordprocessingDocument source = WordprocessingDocument.Open(path, false))
            using (WordprocessingDocument dest = WordprocessingDocument.Open(clonePath, false))
            {
                TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }

        private void CheckSpreadsheetDocument(string path, string clonePath)
        {
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(path, false))
            using (SpreadsheetDocument dest = SpreadsheetDocument.Open(clonePath, false))
            {
                TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }

        private void CheckPresentationDocument(string path, string clonePath)
        {
            using (PresentationDocument source = PresentationDocument.Open(path, false))
            using (PresentationDocument dest = PresentationDocument.Open(clonePath, false))
            {
                TestTools.AssertThatPackagesAreEqual(source, dest);
            }
        }

        [Test]
        public void TestDefaultClone()
        {
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, false))
            using (WordprocessingDocument clone = (WordprocessingDocument)source.Clone())
            {
                Body body = clone.MainDocumentPart.Document.Body;
                body.InsertBefore(new Paragraph(new Run(new Text("Hello World"))), body.FirstChild);
                clone.SaveAs("SaveAndClone\\Default " + documentPath).Close();
            }

            try
            {
                CheckWordprocessingDocument(documentPath, "SaveAndClone\\Default " + documentPath);
            }
            catch (AssertionException)
            {
                // We want the documents to be different.
                return;
            }
            catch (Exception)
            {
                // This is unexpected.
                throw;
            }

            // If the documents are the same, the clone was not writeable, which is an error.
            Assert.Fail();
        }

        [Test]
        public void TestStreamBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("SaveAndClone\\Stream " + documentPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckWordprocessingDocument(documentPath, "SaveAndClone\\Stream " + documentPath);

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("SaveAndClone\\Stream " + spreadsheetPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckSpreadsheetDocument(spreadsheetPath, "SaveAndClone\\Stream " + spreadsheetPath);

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (PresentationDocument dest = (PresentationDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("SaveAndClone\\Stream " + presentationPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckPresentationDocument(presentationPath, "SaveAndClone\\Stream " + presentationPath);
        }

        [Test]
        public void TestFileBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, false))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone("SaveAndClone\\File " + documentPath, false))
            {
                CheckWordprocessingDocument(documentPath, "SaveAndClone\\File " + documentPath);
            }

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, false))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone("SaveAndClone\\File " + spreadsheetPath, false))
            {
                CheckSpreadsheetDocument(spreadsheetPath, "SaveAndClone\\File " + spreadsheetPath);
            }

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, false))
            using (PresentationDocument dest = (PresentationDocument)source.Clone("SaveAndClone\\File " + presentationPath, false))
            {
                CheckPresentationDocument(presentationPath, "SaveAndClone\\File " + presentationPath);
            }
        }

        [Test]
        public void TestPackageBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (Package package = Package.Open("SaveAndClone\\Package " + documentPath, FileMode.Create))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.MainDocumentPart.Document;
            }
            CheckWordprocessingDocument(documentPath, "SaveAndClone\\Package " + documentPath);

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            using (Package package = Package.Open("SaveAndClone\\Package " + spreadsheetPath, FileMode.Create))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.WorkbookPart.Workbook;
            }
            CheckSpreadsheetDocument(spreadsheetPath, "SaveAndClone\\Package " + spreadsheetPath);

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            using (Package package = Package.Open("SaveAndClone\\Package " + presentationPath, FileMode.Create))
            using (PresentationDocument dest = (PresentationDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.PresentationPart.Presentation;
            }
            CheckPresentationDocument(presentationPath, "SaveAndClone\\Package " + presentationPath);
        }

         /// <summary>
        /// Inserts a new paragraph.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="styleId">The style ID or null</param>
        /// <param name="text"></param>
        internal Paragraph InsertParagraph(Body body, string styleId, string text)
        {
            Paragraph p = new Paragraph(new Run(new Text(text)));
            if (styleId != null)
                p.InsertAt(new ParagraphProperties(new ParagraphStyleId { Val = styleId }), 0);
            
            if (body.LastChild != null && body.LastChild is SectionProperties)
                return body.LastChild.InsertBeforeSelf(p);
            else
                return body.AppendChild(p);
        }

        [Test]
        public void TestSave()
        {
            string sourceXml = null;

            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream))
            {
                Document document = dest.MainDocumentPart.Document;
                Body body = document.Body;

                // Make whatever changes you want to make on any part of the document.
                dest.CreateParagraphStyle("MyStyle", "My Test Style", "Normal", "MyStyle");
                InsertParagraph(body, "MyStyle", "Inserted paragraph during TestSave().");

                // Get the document element's XML.
                StringBuilder sb = new StringBuilder();
                using (XmlWriter xw = XmlWriter.Create(sb))
                    document.WriteTo(xw);
                sourceXml = sb.ToString();

                // Save the document. 
                dest.Save();

                // Get the part's root element's XML.
                string partXml = TestTools.GetXmlString(dest.MainDocumentPart);
                Assert.That(sourceXml, Is.EqualTo(partXml));
            }
        }

        [Test]
        public void TestSaveAs()
        {
            // This is probably a bit too much as SaveAs(string) really equals Clone(string).
            // But let's pretend we didn't know that.

            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, false))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.SaveAs("SaveAndClone\\SaveAs " + documentPath))
            {
                CheckWordprocessingDocument(documentPath, "SaveAndClone\\SaveAs " + documentPath);
            }

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, false))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.SaveAs("SaveAndClone\\SaveAs " + spreadsheetPath))
            {
                CheckSpreadsheetDocument(spreadsheetPath, "SaveAndClone\\SaveAs " + spreadsheetPath);
            }

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, false))
            using (PresentationDocument dest = (PresentationDocument)source.SaveAs("SaveAndClone\\SaveAs " + presentationPath))
            {
                CheckPresentationDocument(presentationPath, "SaveAndClone\\SaveAs " + presentationPath);
            }
        }
    }
}
