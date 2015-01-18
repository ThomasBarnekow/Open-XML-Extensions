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
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Presentation;

    [TestFixture]
    public class CreateFromTemplateTests
    {
        static readonly string documentTemplatePath = "Document.dotx";
        static readonly string spreadsheetTemplatePath = "Spreadsheet.xltx";
        static readonly string presentationTemplatePath = "Presentation.potx";

        static readonly string documentPath = "CreateFromTemplate\\Document.docx";
        static readonly string spreadsheetPath = "CreateFromTemplate\\Spreadsheet.xlsx";
        static readonly string presentationPath = "CreateFromTemplate\\Presentation.pptx";

        [TestFixtureSetUp]
        public void SetUp()
        {
            Directory.CreateDirectory("CreateFromTemplate");

            TestTools.RemoveFiles("CreateFromTemplate", "*.docx");
            TestTools.RemoveFiles("CreateFromTemplate", "*.xlsx");
            TestTools.RemoveFiles("CreateFromTemplate", "*.pptx");

            File.Copy(@"..\..\" + documentTemplatePath, documentTemplatePath, true);
            File.Copy(@"..\..\" + spreadsheetTemplatePath, spreadsheetTemplatePath, true);
            File.Copy(@"..\..\" + presentationTemplatePath, presentationTemplatePath, true);
        }

        [Test]
        public void TestWordprocessingDocument()
        {
            using (WordprocessingDocument packageDocument = WordprocessingDocument.CreateFromTemplate(documentTemplatePath))
            {
                MainDocumentPart part = packageDocument.MainDocumentPart;
                Document root = part.Document;

                packageDocument.SaveAs(documentPath).Close();

                // We are fine if we have not run into an exception.
                Assert.True(true);
            }
        }

        [Test]
        public void TestSpreadsheetDocument()
        {
            using (SpreadsheetDocument packageDocument = SpreadsheetDocument.CreateFromTemplate(spreadsheetTemplatePath))
            {
                WorkbookPart part = packageDocument.WorkbookPart;
                Workbook root = part.Workbook;

                packageDocument.SaveAs(spreadsheetPath).Close();

                // We are fine if we have not run into an exception.
                Assert.True(true);
            }
        }

        [Test]
        public void TestPresentationDocument()
        {
            using (PresentationDocument packageDocument = PresentationDocument.CreateFromTemplate(presentationTemplatePath))
            {
                PresentationPart part = packageDocument.PresentationPart;
                Presentation root = part.Presentation;

                packageDocument.SaveAs(presentationPath).Close();

                // We are fine if we have not run into an exception.
                Assert.True(true);
            }
        }
    }
}
