/*
 * WordprocessingMemoryStream.cs - MemoryStream for WordprocessingDocuments
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
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.IO
{
    /// <summary>
    /// A <see cref="MemoryStream"/> for <see cref="WordprocessingDocument"/>s 
    /// with a number of specific helper methods.
    /// </summary>
    public class WordprocessingMemoryStream : OpenXmlMemoryStream
    {
        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class 
        /// with an expandable capacity initialized to zero.
        /// </summary>
        protected WordprocessingMemoryStream()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class 
        /// with an expandable capacity and the contents of the given buffer.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public WordprocessingMemoryStream(byte[] buffer)
            : base(buffer)
        { }

        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class with an 
        /// expandable capacity and the contents of the given buffer. 
        /// Will use the given path when saving the stream.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        /// <param name="path">The path to be used when saving the stream</param>
        public WordprocessingMemoryStream(byte[] buffer, string path)
            : base(buffer, path)
        { }

        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class 
        /// with an expandable capacity and the contents of the given file.
        /// </summary>
        /// <param name="path"></param>
        public WordprocessingMemoryStream(string path)
            : base(path)
        { }

        /// <summary>
        /// Initializes a new instance of the WordprocessingMemoryStream class 
        /// with an expandable capacity and the contents of the given stream.
        /// </summary>
        /// <param name="stream"></param>
        public WordprocessingMemoryStream(Stream stream)
            : base(stream)
        { }

        /// <summary>
        /// Initializes this stream from a byte array.
        /// </summary>
        /// <param name="buffer">The byte array.</param>
        protected override void InitFromByteArray(byte[] buffer)
        {
            base.InitFromByteArray(buffer);
            if (DocumentType != typeof(WordprocessingDocument))
                throw new ArgumentException("Not a WordprocessingDocument: " + DocumentType);
        }

        /// <summary>
        /// Creates a new <see cref="WordprocessingMemoryStream"/> containing a
        /// <see cref="WordprocessingDocument"/> with a "minimum document", i.e., 
        /// one having a MainDocumentPart with a w:document element containing 
        /// an empty w:body. 
        /// </summary>
        /// <returns></returns>
        public static WordprocessingMemoryStream Create()
        {
            WordprocessingMemoryStream stream = new WordprocessingMemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Create(
                stream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart part = doc.AddMainDocumentPart();
                part.Document = new Document(new Body());
            }
            return stream;
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="WordprocessingDocument"/>
        ///  class from this stream. 
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public override OpenXmlPackage OpenDocument(bool isEditable, OpenSettings openSettings)
        {
            return OpenWordprocessingDocument(isEditable, openSettings);
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="WordprocessingDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public WordprocessingDocument OpenWordprocessingDocument(bool isEditable)
        {
            return OpenWordprocessingDocument(isEditable, new OpenSettings());
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="WordprocessingDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public WordprocessingDocument OpenWordprocessingDocument(bool isEditable, OpenSettings openSettings)
        {
            return WordprocessingDocument.Open(this, isEditable, openSettings);
        }
    }
}
