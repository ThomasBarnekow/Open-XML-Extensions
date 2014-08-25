/*
 * SpreadsheetMemoryStream.cs - MemoryStream for SpreadsheetDocuments
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

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.IO
{
    /// <summary>
    /// A <see cref="MemoryStream"/> for <see cref="SpreadsheetDocument"/>s 
    /// with a number of specific helper methods.
    /// </summary>
    public class SpreadsheetMemoryStream : OpenXmlMemoryStream
    {
        /// <summary>
        /// Initializes a new instance of the SpreadsheetMemoryStream class with
        /// an expandable capacity initialized to zero.
        /// </summary>
        protected SpreadsheetMemoryStream()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of the SpreadsheetMemoryStream class with 
        /// an expandable capacity and the contents of the given buffer.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        public SpreadsheetMemoryStream(byte[] buffer)
            : base(buffer)
        { }

        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class with an 
        /// expandable capacity and the contents of the given buffer. 
        /// Will use the given path when saving the stream.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        /// <param name="path">The path to be used when saving the stream</param>
        public SpreadsheetMemoryStream(byte[] buffer, string path)
            : base(buffer, path)
        { }

        /// <summary>
        /// Initializes a new instance of the SpreadsheetMemoryStream class with 
        /// an expandable capacity and the contents of the given file.
        /// </summary>
        /// <param name="path"></param>
        public SpreadsheetMemoryStream(string path)
            : base(path)
        { }

        /// <summary>
        /// Initializes a new instance of the SpreadsheetMemoryStream class with 
        /// an expandable capacity and the contents of the given stream.
        /// </summary>
        /// <param name="stream"></param>
        public SpreadsheetMemoryStream(Stream stream)
            : base(stream)
        { }

        /// <summary>
        /// Initializes this stream from a byte array.
        /// </summary>
        /// <param name="buffer">The byte array.</param>
        protected override void InitFromByteArray(byte[] buffer)
        {
            base.InitFromByteArray(buffer);
            if (DocumentType != typeof(SpreadsheetDocument))
                throw new ArgumentException("Not a SpreadsheetDocument: " + DocumentType);
        }

        /// <summary>
        /// Creates a new <see cref="SpreadsheetMemoryStream"/> containing a
        /// "minimum workbook" as defined by the Standard ECMA-376, i.e., 
        /// an x:workbook with an x:sheets child containing a single x:sheet
        /// with a sheet ID and a relationship ID that points to the location 
        /// of the sheet definition. The latter is the root element of a 
        /// WorksheetPart and contains an empty x:sheetData element.
        /// </summary>
        /// <returns></returns>
        public static SpreadsheetMemoryStream Create()
        {
            SpreadsheetMemoryStream stream = new SpreadsheetMemoryStream();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                // When creating a new SpreadsheetDocument, we will simply define
                // our own relationship IDs.
                Sheet sheet = new Sheet { Name = "Sheet1", SheetId = 1, Id = "rId1" };

                // Create a WorkbookPart with an x:workbook root element containing
                // a single sheet reference.
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets(sheet));

                // Create a WorksheetPart with an x:worksheet root element containing
                // an empty x:sheetData child.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheet.Id);
                worksheetPart.Worksheet = new Worksheet(new SheetData());
            }
            return stream;
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="SpreadsheetDocument"/>
        ///  class from this stream. 
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public override OpenXmlPackage OpenDocument(bool isEditable, OpenSettings openSettings)
        {
            return OpenSpreadsheetDocument(isEditable, openSettings);
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="SpreadsheetDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public SpreadsheetDocument OpenSpreadsheetDocument(bool isEditable)
        {
            return OpenSpreadsheetDocument(isEditable, new OpenSettings());
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="SpreadsheetDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public SpreadsheetDocument OpenSpreadsheetDocument(bool isEditable, OpenSettings openSettings)
        {
            return SpreadsheetDocument.Open(this, isEditable, openSettings);
        }
    }
}
