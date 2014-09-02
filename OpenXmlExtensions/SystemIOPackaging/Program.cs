/*
 * Program.cs - Test driver for FlatOpcPackage
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

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SystemIOPackaging
{
    /// <summary>
    /// This class tests the FlatOpcPackage functionality. It will be replaced
    /// by NUnit-based tests sooner or later.
    /// </summary>
    class Program
    {
        static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        static void Main(string[] args)
        {
            // SimpleFlatOpcPackageTest();
            WordprocessingDocumentBasedTest();

            Console.WriteLine("\nHit any key to continue...");
            Console.ReadKey();
        }

        static void SimpleFlatOpcPackageTest()
        {
            // XDocument doc = XDocument.Load("Hello World.xml");
            // FlatOpcPackage package = new FlatOpcPackage(doc);

            // FileStream fileStream = new FileStream("Hello World.xml", FileMode.Open, FileAccess.ReadWrite);
            // FlatOpcPackage package = FlatOpcPackage.Open(fileStream, FileMode.Open, FileAccess.ReadWrite);

            FlatOpcPackage package = FlatOpcPackage.Open("Hello World.xml", FileMode.Open, FileAccess.ReadWrite);

            foreach (PackagePart part in package.GetParts())
                Console.WriteLine(part.Uri);

            PackagePart pp = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            XDocument main;
            using (Stream stream = pp.GetStream(FileMode.Open, FileAccess.Read))
            {
                main = XDocument.Load(stream);
                Console.WriteLine("\nWriting document before making changes ...");
                Console.WriteLine(main.Declaration.ToString());
                Console.WriteLine(main.ToString());
            }

            // Create part contents based on what we've read before, introducing
            // some changes.
            using (Stream stream = pp.GetStream(FileMode.Create, FileAccess.ReadWrite))
            {
                XElement body = main.Root.Descendants(w + "body").First();
                body.AddFirst(
                    new XElement(w + "p",
                        new XElement(w + "r",
                            new XElement(w + "t", "This was inserted into the part"))));

                main.Save(stream);
            }

            // Read what we've written.
            using (Stream stream = pp.GetStream(FileMode.Open, FileAccess.Read))
            {
                main = XDocument.Load(stream);
                Console.WriteLine("\nWriting document after making changes ...");
                Console.WriteLine(main.Declaration.ToString());
                Console.WriteLine(main.ToString());
            }

            package.Flush();

            // Get the new document.
            XDocument newDoc = package.Document;
        }

        static void WordprocessingDocumentBasedTest()
        {
            // Create a new package.
            FlatOpcPackage package = FlatOpcPackage.Open("Test.xml", FileMode.Create);
            using (WordprocessingDocument doc = WordprocessingDocument.Create(package, 
                WordprocessingDocumentType.Document))
            {
                MainDocumentPart part = doc.AddMainDocumentPart();
                part.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Hello World!")))));

            }
            
            // Open that package again.
            package = FlatOpcPackage.Open("Test.xml", FileMode.Open);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(package))
            {
                Document document = doc.MainDocumentPart.Document;

                // Write main document part's contents.
                Console.WriteLine("\nHere's what we've created from scratch:");
                StringBuilder sb = new StringBuilder();
                using (XmlWriter xw = XmlWriter.Create(sb, new XmlWriterSettings { Indent = true }))
                {
                    document.WriteTo(xw);
                }
                Console.WriteLine(sb.ToString());
            }

            // Let's do something bigger, i.e., clone an OPC package-based Word document,
            // to a FlatOpcPackage-based Open XML document.
            package = FlatOpcPackage.Open("Clone.xml", FileMode.Create);
            using (WordprocessingDocument original = WordprocessingDocument.Open("Document.docx", false))
            using (WordprocessingDocument clone = (WordprocessingDocument)original.Clone(package))
            {
                Document document = clone.MainDocumentPart.Document;

                // Write main document part's contents.
                Console.WriteLine("\nHere's what we've just cloned:");
                StringBuilder sb = new StringBuilder();
                using (XmlWriter xw = XmlWriter.Create(sb, new XmlWriterSettings { Indent = true }))
                {
                    document.WriteTo(xw);
                }
                Console.WriteLine(sb.ToString());
            }

            // Let's now open the clone from an XDocument and look at the contents.
            XDocument cloneDoc = XDocument.Load("Clone.xml");
            package = FlatOpcPackage.Open(cloneDoc);
            using (WordprocessingDocument clone = WordprocessingDocument.Open(package))
            {
                Document document = clone.MainDocumentPart.Document;
                Paragraph p = document.Body.Elements<Paragraph>().First();

                Console.WriteLine("\nHere's the first run of text of the XDocument we've opened as a FlatOpcPackage:");
                Console.WriteLine(p.Descendants<Text>().First().Text);

                p.InsertBeforeSelf(
                    new Paragraph(
                        new Run(
                            new Text("Inserted before first paragraph"))));
            }

            // Let's access the package's Document again to see whether that works.
            XDocument testDoc = package.Document;
            XElement testElement = testDoc.Descendants(w + "t").First();
            
            Console.WriteLine("\nHere's what we got from the XDocument after closing the WordprocessingDocument:");
            Console.WriteLine(testElement.Value);
        }
    }
}
