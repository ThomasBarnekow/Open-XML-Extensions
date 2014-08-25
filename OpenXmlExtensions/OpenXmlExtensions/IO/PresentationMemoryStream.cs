/*
 * PresentationMemoryStream.cs - MemoryStream for PresentationDocuments
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
using DocumentFormat.OpenXml.Presentation;

namespace DocumentFormat.OpenXml.IO
{
    /// <summary>
    /// A <see cref="MemoryStream"/> for <see cref="PresentationDocument"/>s 
    /// with a number of specific helper methods.
    /// </summary>
    public class PresentationMemoryStream : OpenXmlMemoryStream
    {
        /// <summary>
        /// Initializes a new instance of the PresentationMemoryStream class 
        /// with an expandable capacity initialized to zero.
        /// </summary>
        protected PresentationMemoryStream()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of the PresentationMemoryStream class 
        /// with an expandable capacity and the contents of the given buffer.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public PresentationMemoryStream(byte[] buffer)
            : base(buffer)
        { }

        /// <summary>
        /// Initializes a new instance of the PresentationMemoryStream class with an 
        /// expandable capacity and the contents of the given buffer. 
        /// Will use the given path when saving the stream.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        /// <param name="path">The path to be used when saving the stream</param>
        public PresentationMemoryStream(byte[] buffer, string path)
            : base(buffer, path)
        { }

        /// <summary>
        /// Initializes a new instance of the PresentationMemoryStream class 
        /// with an expandable capacity and the contents of the given file.
        /// </summary>
        /// <param name="path"></param>
        public PresentationMemoryStream(string path)
            : base(path)
        { }

        /// <summary>
        /// Initializes a new instance of the PresentationMemoryStream class 
        /// with an expandable capacity and the contents of the given stream.
        /// </summary>
        /// <param name="stream"></param>
        public PresentationMemoryStream(Stream stream)
            : base(stream)
        { }

        /// <summary>
        /// Initializes this stream from a byte array.
        /// </summary>
        /// <param name="buffer">The byte array.</param>
        protected override void InitFromByteArray(byte[] buffer)
        {
            base.InitFromByteArray(buffer);
            if (DocumentType != typeof(PresentationDocument))
                throw new ArgumentException("Not a PresentationDocument: " + DocumentType);
        }

        /// <summary>
        /// Creates a new <see cref="PresentationMemoryStream"/> containing a
        /// <see cref="PresentationDocument"/> with a "minimum document".
        /// </summary>
        /// <returns></returns>
        public static PresentationMemoryStream Create()
        {
            PresentationMemoryStream stream = new PresentationMemoryStream();
            using (PresentationDocument doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
            {
                PresentationPart part = doc.AddPresentationPart();
                part.Presentation = new Presentation.Presentation(
                    new SlideMasterIdList(),
                    new SlideIdList(),
                    new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
                    new NotesSize { Cx = 6858000, Cy = 9144000 });
            }
            return stream;
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="PresentationDocument"/>
        ///  class from this stream. 
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public override OpenXmlPackage OpenDocument(bool isEditable, OpenSettings openSettings)
        {
            return OpenPresentationDocument(isEditable, openSettings);
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="PresentationDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public PresentationDocument OpenPresentationDocument(bool isEditable)
        {
            return OpenPresentationDocument(isEditable, new OpenSettings());
        }

        /// <summary>
        ///  Creates a new instance of the <see cref="PresentationDocument"/>
        ///  class from this stream.
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public PresentationDocument OpenPresentationDocument(bool isEditable, OpenSettings openSettings)
        {
            return PresentationDocument.Open(this, isEditable, openSettings);
        }
    }
}
