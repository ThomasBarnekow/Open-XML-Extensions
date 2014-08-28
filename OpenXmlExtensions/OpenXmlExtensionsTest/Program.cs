using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Transforms;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlExtensionsTest
{
    class Program
    {
        static string documentPath = "Document.docx";
        static string spreadsheetPath = "Spreadsheet.xlsx";
        static string presentationPath = "Presentation.pptx";
        
        static void Main(string[] args)
        {
            TestStreamBasedClone();
            TestFileBasedClone();
            TestPackageBasedClone();

            TestSave();
            TestSaveAs();
        }

        static void CheckWordprocessingDocument(string path)
        {
            using (WordprocessingDocument dest = WordprocessingDocument.Open(path, false))
            {
                OpenXmlElement root = dest.MainDocumentPart.Document;
            }
        }

        static void CheckSpreadsheetDocument(string path)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
            {
                OpenXmlElement root = doc.WorkbookPart.Workbook;
            }
        }

        static void CheckPresentationDocument(string path)
        {
            using (PresentationDocument doc = PresentationDocument.Open(path, false))
            {
                OpenXmlElement root = doc.PresentationPart.Presentation;
            }
        }

        static void TestStreamBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("Stream " + documentPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckWordprocessingDocument("Stream " + documentPath);

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("Stream " + spreadsheetPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckSpreadsheetDocument("Stream " + spreadsheetPath);

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            using (MemoryStream memoryStream = new MemoryStream())
            using (PresentationDocument dest = (PresentationDocument)source.Clone(memoryStream, true))
            using (FileStream fileStream = new FileStream("Stream " + presentationPath, FileMode.Create))
            {
                memoryStream.WriteTo(fileStream);
            }
            CheckPresentationDocument("Stream " + presentationPath);
        }

        static void TestFileBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone("File " + documentPath, false))
            {
                CheckWordprocessingDocument("File " + documentPath);
            }

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone("File " + spreadsheetPath, false))
            {
                CheckSpreadsheetDocument("File " + spreadsheetPath);
            }

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            using (PresentationDocument dest = (PresentationDocument)source.Clone("File " + presentationPath, false))
            {
                CheckPresentationDocument("File " + presentationPath);
            }
        }

        static void TestPackageBasedClone()
        {
            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (Package package = Package.Open("Package " + documentPath, FileMode.Create))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.MainDocumentPart.Document;
            }
            CheckWordprocessingDocument("Package " + documentPath);

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, true))
            using (Package package = Package.Open("Package " + spreadsheetPath, FileMode.Create))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.WorkbookPart.Workbook;
            }
            CheckSpreadsheetDocument("Package " + spreadsheetPath);

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, true))
            using (Package package = Package.Open("Package " + presentationPath, FileMode.Create))
            using (PresentationDocument dest = (PresentationDocument)source.Clone(package))
            {
                OpenXmlElement root = dest.PresentationPart.Presentation;
            }
            CheckPresentationDocument("Package " + presentationPath);
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
            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, true))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.Clone(memoryStream))
            {
                Document document = dest.MainDocumentPart.Document;
                Body body = document.Body;

                // Make whatever changes you want to make on any part of the document.
                dest.CreateParagraphStyle("MyStyle", "My Test Style", "Normal", "MyStyle");
                InsertParagraph(body, "MyStyle", "Inserted paragraph during TestSave().");

                // Save the document. 
                // If we knew exactly what we changed, we could also do it like this:
                //     document.Save();
                //     dest.Package.Flush();
                // However, we'd have to save each root element, e.g., Document, that
                // changed. Save will do the job for us and also flush the Package.
                dest.Save();
                
                // Now, let's see whether we can save the MemoryStream to a file.
                using (FileStream fileStream = new FileStream("Save " + documentPath, FileMode.Create))
                    memoryStream.WriteTo(fileStream);
            }
            CheckWordprocessingDocument("Save " + documentPath);
        }

        static void TestSaveAs()
        {
            // This is probably a bit too much as SaveAs(string) really equals Clone(string).
            // But let's pretend we didn't know that.

            // Test WordprocessingDocument.
            using (WordprocessingDocument source = WordprocessingDocument.Open(documentPath, false))
            using (WordprocessingDocument dest = (WordprocessingDocument)source.SaveAs("SaveAs " + documentPath))
            {
                CheckWordprocessingDocument("SaveAs " + documentPath);
            }

            // Test SpreadsheetDocument.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(spreadsheetPath, false))
            using (SpreadsheetDocument dest = (SpreadsheetDocument)source.SaveAs("SaveAs " + spreadsheetPath))
            {
                CheckSpreadsheetDocument("SaveAs " + spreadsheetPath);
            }

            // Test PresentationDocument.
            using (PresentationDocument source = PresentationDocument.Open(presentationPath, false))
            using (PresentationDocument dest = (PresentationDocument)source.SaveAs("SaveAs " + presentationPath))
            {
                CheckPresentationDocument("SaveAs " + presentationPath);
            }
        }
    }
}
