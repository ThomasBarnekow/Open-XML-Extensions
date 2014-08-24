/*
 * Program.cs - Various tests for Open XML MemoryStreams
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
using System.IO.Packaging;
using System.Text;
using System.Xml;

using DocumentFormat.OpenXml.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using OpenXmlPowerTools;

namespace OpenXmlSdkTest
{
    /// <summary>
    /// Class implementing various tests, however, with a focus on the Wordprocessing
    /// side of things. 
    /// </summary>
    class Program
    {
        /// <summary>
        /// Runs various tests.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                // Perform a test with the OpenXmlMemoryStreamDocument provided
                // by the PowerTools for Open XML.
                TestOpenXmlMemoryStreamDocument_Failing();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                Console.WriteLine();
                Console.WriteLine("Press any key to continue ...");
                Console.ReadKey();
            }

            try
            {
                TestWordprocessingMemoryStream();
                TestSpreadsheetMemoryStream();
                TestPresentationMemoryStream();

                SimpleWordprocessingApplication();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                Console.WriteLine();
                Console.WriteLine("Press any key to continue ...");
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Inserts a new paragraph.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="text"></param>
        static void InsertParagraph(Body body, string text)
        {
            Paragraph p = new Paragraph(new Run(new Text(text)));
            if (body.LastChild != null && body.LastChild is SectionProperties)
                body.LastChild.InsertBeforeSelf(p);
            else
                body.Append(p);
        }

        /// <summary>
        /// This test essentially shows what doesn't work with the
        /// OpenXmlMemoryStreamDocument. The issue is that there is no way to
        /// reuse the MemoryStream contained in the OpenXmlMemoryStreamDocument.
        /// </summary>
        static void TestOpenXmlMemoryStreamDocument_Failing()
        {
            Console.WriteLine("\nTest OpenXmlMemoryStreamDocument\n");

            WmlDocument wmlDoc = new WmlDocument(@"Document.docx");
            OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(wmlDoc);
            
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                Document document = wordDoc.MainDocumentPart.Document;
                InsertParagraph(document.Body, "This is the first paragraph.");
                document.Save();
            }

            // This second pass will fail with an Exception, saying the Package has
            // already been closed.
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                Document document = wordDoc.MainDocumentPart.Document;
                InsertParagraph(document.Body, "This is the second paragraph.");
                document.Save();
            }

            StringBuilder sb = new StringBuilder();
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            using (XmlWriter xw = XmlWriter.Create(sb, new XmlWriterSettings { Indent = true }))
            {
                Document document = wordDoc.MainDocumentPart.Document;
                document.WriteTo(xw);
            }
            Console.Write(sb);

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        /// <summary>
        /// Tests the <see cref="WordprocessingMemoryStream"/>.
        /// </summary>
        static void TestWordprocessingMemoryStream()
        {
            Console.WriteLine("\nTest WordprocessingMemoryStream\n");

            // First alternative: Initialize stream from existing file.
            // WordprocessingMemoryStream stream = new WordprocessingMemoryStream(@"..\..\Hello World.docx");
            
            // Other alternative: Use WmlDocument to initialize stream.
            // WmlDocument wmlDoc1 = new WmlDocument(@"Document.docx");
            // WordprocessingMemoryStream stream = PtMemoryStreamTools.CreateWordprocessingMemoryStream(wmlDoc1);

            // Other alternative: Create stream containing empty WordprocessingDocument.
            WordprocessingMemoryStream stream = WordprocessingMemoryStream.Create();

            using (WordprocessingDocument wordDoc = stream.OpenWordprocessingDocument(true))
            {
                Document document = wordDoc.MainDocumentPart.Document;
                InsertParagraph(document.Body, "This is the first paragraph.");
                document.Save();
            }

            using (WordprocessingDocument wordDoc = stream.OpenWordprocessingDocument(true))
            {
                Document document = wordDoc.MainDocumentPart.Document;
                InsertParagraph(document.Body, "This is the second paragraph.");
                document.Save();
            }

            // Print MainDocumentPart.
            Console.WriteLine("\nPrinting MainDocumentPart");
            StringBuilder sb = new StringBuilder();
            using (WordprocessingDocument wordDoc = stream.OpenWordprocessingDocument(false))
            using (XmlWriter xw = XmlWriter.Create(sb, new XmlWriterSettings { Indent = true }))
            {
                Document document = wordDoc.MainDocumentPart.Document;
                document.WriteTo(xw);
            }
            Console.WriteLine(sb.ToString());

            // Check Package.
            Console.WriteLine("\nPrinting PackagePart URIs");
            using (Package package = stream.OpenPackage())
            {
                foreach (PackagePart part in package.GetParts())
                    Console.WriteLine(part.Uri.ToString());
            }

            // Save stream.
            stream.SaveAs(@"Document (Saved by Stream).docx");

            stream.Close();
            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        /// <summary>
        /// Test the <see cref="SpreadsheetMemoryStream"/>.
        /// </summary>
        static void TestSpreadsheetMemoryStream()
        {
            Console.WriteLine("\nTest SpreadsheetMemoryStream\n");

            // First alternative: Initialize stream from existing file.
            // SpreadsheetMemoryStream stream = new SpreadsheetMemoryStream(@"Spreadsheet.xlsx");
            
            // Alternative: Create new workbook
            SpreadsheetMemoryStream stream = SpreadsheetMemoryStream.Create();

            // Save stream.
            stream.SaveAs(@"Spreadsheet (Saved by Stream).xlsx");

            stream.Close();
            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        /// <summary>
        /// Tests the <see cref="PresentationMemoryStream"/>.
        /// </summary>
        static void TestPresentationMemoryStream()
        {
            Console.WriteLine("\nTest PresentationMemoryStream\n");

            // First alternative: Initialize stream from existing file.
            // PresentationMemoryStream stream = new PresentationMemoryStream(@"Presentation.pptx");

            // Alternative: Create new presentation.
            PresentationMemoryStream stream = PresentationMemoryStream.Create();

            // Save stream.
            stream.SaveAs(@"Presentation (Saved by Stream).pptx");

            stream.Close();
            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        static void SimpleWordprocessingApplication()
        {
            // Generate a Word document in multiple processing steps based on 
            // a "minimum document" created by WordprocessingMemoryStream.
            using (WordprocessingMemoryStream stream = WordprocessingMemoryStream.Create())
            {
                // Perform a second processing step. Do whatever you like with the
                // WordprocessingDocument. This example just inserts a paragraph.
                // When leaving the scope of the using statement, the stream will
                // contain a perfectly fine WordprocessingDocument which we can
                // continue to process in further steps.
                using (WordprocessingDocument wordDoc = stream.OpenWordprocessingDocument(true))
                {
                    Document document = wordDoc.MainDocumentPart.Document;
                    InsertParagraph(document.Body, "This is the first paragraph.");
                }

                // Perform a second processing step, using the same WordprocessingMemoryStream.
                // Again, do whatever you like with the WordprocessingDocument. This example
                // just creates another paragraph. When leaving the using statement, the
                // WordprocessingDocument will be closed, leaving us with a stream that
                // can be reused over and over again without copying any data.
                using (WordprocessingDocument wordDoc = stream.OpenWordprocessingDocument(true))
                {
                    Document document = wordDoc.MainDocumentPart.Document;
                    InsertParagraph(document.Body, "This is the second paragraph.");
                }

                // Lastly, let's save the stream contents to a file.
                stream.SaveAs("Generated Document.docx");
            }
        }
    }
}
