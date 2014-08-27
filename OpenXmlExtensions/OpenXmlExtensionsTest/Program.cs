using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Transforms;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlExtensionsTest
{
    class Program
    {
        static void Main(string[] args)
        {
            TestStreamBasedClone();
            TestPackageOnFileBasedClone();
            TestPackageOnStreamBasedClone();
            TestFileBasedClone();

            TestSave();
        }

        static void TestStreamBasedClone()
        {
            Console.WriteLine("\nTesting MemoryStream-based cloning ...");

            string sourcePath = "Document.docx";
            string destPath = "StreamBasedClone.docx";

            // First pass: Create a clone on a MemoryStream and write the latter to a file
            // right away. This tests whether the MemoryStream contains a valid document
            // after having cloned the document.
            using (WordprocessingDocument source = WordprocessingDocument.Open(sourcePath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (FileStream fileStream = new FileStream(destPath, FileMode.Create))
            {
                WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream, true);
                Console.WriteLine("\nListing all parts after creating the clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);

                memoryStream.WriteTo(fileStream);
            }

            // Second pass: Open the file saved from the MemoryStream containing the clone.
            // Perform some operations to see whether there are issues.
            using (WordprocessingDocument dest = WordprocessingDocument.Open(destPath, true))
            {
                Console.WriteLine("\nListing all parts after reopening the saved clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        static void TestPackageOnFileBasedClone()
        {
            Console.WriteLine("\nTesting Package on File-based cloning ...");

            string sourcePath = "Document.docx";
            string destPath = "PackageOnFileBasedClone.docx";

            // First pass: Create a clone using the package-based Clone method. This should
            // automatically save the file (when leaving the scope of the using statement).
            using (WordprocessingDocument source = WordprocessingDocument.Open(sourcePath, true))
            using (Package package = Package.Open(destPath, FileMode.Create))
            {
                WordprocessingDocument dest = (WordprocessingDocument)source.Clone(package);
                Console.WriteLine("\nListing all parts after creating the clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
                
                XDocument doc = FlatOpcTransform.ToFlatOpc(dest);
                using (XmlWriter xw = XmlWriter.Create("PackageOnFileBasedClone.xml", new XmlWriterSettings { Indent = true }))
                    doc.WriteTo(xw);
            }

            // Second pass: Open the file created by the Package.Open(string, FileMode) method.
            // Perform some operations to see whether there are issues.
            using (WordprocessingDocument dest = WordprocessingDocument.Open(destPath, true))
            {
                Console.WriteLine("\nListing all parts after reopening the saved clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        static void TestPackageOnStreamBasedClone()
        {
            Console.WriteLine("\nTesting Package on MemoryStream-based cloning ...");

            string sourcePath = "Document.docx";
            string destPath = "PackageOnStreamBasedClone.docx";

            // First pass: Clone a document, using a combination of MemoryStream and Package.
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (Package package = Package.Open(memoryStream, FileMode.Create))
                using (WordprocessingDocument source = WordprocessingDocument.Open(sourcePath, true))
                using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(package))
                {
                    Console.WriteLine("\nListing all parts after creating the clone:");
                    foreach (var part in dest.GetAllParts())
                        Console.WriteLine(part.Uri);

                    XDocument doc = FlatOpcTransform.ToFlatOpc(dest);
                    using (XmlWriter xw = XmlWriter.Create("PackageOnStreamBasedClone.xml", new XmlWriterSettings { Indent = true }))
                        doc.WriteTo(xw);
                }

                using (FileStream fileStream = new FileStream(destPath, FileMode.Create))
                    memoryStream.WriteTo(fileStream);
            }

            // Second pass: Open the file saved from the MemoryStream containing the clone.
            // Perform some operations to see whether there are issues.
            using (WordprocessingDocument dest = WordprocessingDocument.Open(destPath, true))
            {
                Console.WriteLine("\nListing all parts after reopening the saved clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        static void TestFileBasedClone()
        {
            Console.WriteLine("\nTesting File-based cloning ...");

            string sourcePath = "Document.docx";
            string destPath = "FileBasedClone.docx";

            using (WordprocessingDocument source = WordprocessingDocument.Open(sourcePath, true))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(destPath))
            {
                Console.WriteLine("\nListing all parts after creating the clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            // Second pass: Open the file saved from the MemoryStream containing the clone.
            // Perform some operations to see whether there are issues.
            using (WordprocessingDocument dest = WordprocessingDocument.Open(destPath, true))
            {
                Console.WriteLine("\nListing all parts after reopening the saved clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }

        /// <summary>
        /// Inserts a new paragraph.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="styleId">The style ID or null</param>
        /// <param name="text"></param>
        static Paragraph InsertParagraph(Body body, string styleId, string text)
        {
            Paragraph p = new Paragraph(new Run(new Text(text)));
            if (styleId != null)
                p.InsertAt(new ParagraphProperties(new ParagraphStyleId { Val = styleId }), 0);
            
            if (body.LastChild != null && body.LastChild is SectionProperties)
                return body.LastChild.InsertBeforeSelf(p);
            else
                return body.AppendChild(p);
        }

        static void TestSave()
        {
            Console.WriteLine("\nTesting Save ...");

            string sourcePath = "Document.docx";
            string destPath = "SavedDocument.docx";

            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument source = WordprocessingDocument.Open(sourcePath, true))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream))
            {
                Document document = dest.MainDocumentPart.Document;
                Body body = document.Body;

                // Make whatever changes you want to make on any part of the document.
                dest.CreateParagraphStyle("MyStyle", "My Test Style", "Normal", "MyStyle");
                InsertParagraph(body, "MyStyle", "Inserted paragraph");

                // Save the document. 
                // If we knew exactly what we changed, we could also do it like this:
                //     document.Save();
                //     dest.Package.Flush();
                // However, we'd have to save each root element, e.g., Document, that
                // changed. Save will do the job for us and also flush the Package.
                dest.Save();
                
                // Now, let's see whether we can save the MemoryStream to a file.
                using (FileStream fileStream = new FileStream(destPath, FileMode.Create))
                    memoryStream.WriteTo(fileStream);
            }

            // Second pass: Open the file saved from the MemoryStream containing the clone.
            // Perform some operations to see whether there are issues.
            using (WordprocessingDocument dest = WordprocessingDocument.Open(destPath, true))
            {
                Console.WriteLine("\nListing all parts after reopening the saved clone:");
                foreach (var part in dest.GetAllParts())
                    Console.WriteLine(part.Uri);
            }

            Console.WriteLine("\nPress any key to continue ...");
            Console.ReadKey();
        }
    }
}
