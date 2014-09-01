using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Linq;

using System.IO;
using System.IO.Packaging;
using System.IO.Packaging.FlatOpc;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SystemIOPackaging
{
    class Program
    {
        static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        static void Main(string[] args)
        {
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
            FlatOpcPackage package = FlatOpcPackage.Open("Test.xml", FileMode.OpenOrCreate);

            using (WordprocessingDocument doc = WordprocessingDocument.Create(package, 
                WordprocessingDocumentType.Document))
            {
                MainDocumentPart part = doc.AddMainDocumentPart();
                part.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Hello World!")))));

                // Not sure why we have to call this directly.
                // We apparently have to do this, at least for now.
                // There's probably a bug.
                part.Document.Save();
                package.Flush();
            }
        }
    }
}
