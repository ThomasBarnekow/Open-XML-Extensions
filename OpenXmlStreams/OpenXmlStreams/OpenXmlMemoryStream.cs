/*
 * OpenXmlMemoryStream.cs - MemoryStream for Open XML Documents
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
using System.Linq;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.IO
{
    /// <summary>
    /// This class implements a <see cref="MemoryStream"/> used to store 
    /// <see cref="OpenXmlPackage"/>s and serves as the base class for more 
    /// concrete MemoryStreams containing <see cref="WordprocessingDocument"/>s,
    /// <see cref="SpreadsheetDocument"/>s, and <see cref="PresentationDocument"/>s.
    /// </summary>
    public abstract class OpenXmlMemoryStream : MemoryStream
    {
        #region Constructors and Initializers

        /// <summary>
        /// Initializes a new instance of the OpenXmlMemoryStream class with an 
        /// expandable capacity initialized to zero.
        /// </summary>
        protected OpenXmlMemoryStream()
            : base()
        { }

        /// <summary>
        /// Initializes a new instance of the OpenXmlMemoryStream class with an 
        /// expandable capacity and the contents of the given buffer.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        protected OpenXmlMemoryStream(byte[] buffer)
            : base()
        {
            InitFromByteArray(buffer);
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlMemoryStream class with an 
        /// expandable capacity and the contents of the given buffer. 
        /// Will use the given path when saving the stream using <see cref="Save"/>.
        /// </summary>
        /// <param name="buffer">The buffer</param>
        /// <param name="path">The path to be used when saving the stream</param>
        protected OpenXmlMemoryStream(byte[] buffer, string path)
            : base()
        {
            InitFromByteArray(buffer);
            Path = path;
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlMemoryStream class with an 
        /// expandable capacity and the contents of the given file.
        /// </summary>
        /// <param name="path"></param>
        protected OpenXmlMemoryStream(string path)
            : base()
        {
            if (path == null)
                throw new ArgumentNullException("path");

            InitFromByteArray(File.ReadAllBytes(path));
            Path = path;
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlMemoryStream class with an 
        /// expandable capacity and the contents of the given stream.
        /// </summary>
        /// <param name="stream"></param>
        protected OpenXmlMemoryStream(Stream stream)
            : base()
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, (int)stream.Length);
            InitFromByteArray(buffer);
        }

        /// <summary>
        /// Initializes this stream from a byte array.
        /// </summary>
        /// <param name="buffer">The byte array</param>
        protected virtual void InitFromByteArray(byte[] buffer)
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            this.Write(buffer, 0, buffer.Length);

            // Determine document type.
            using (Package package = Package.Open(this, FileMode.Open, FileAccess.Read))
            {
                PackageRelationship relationship = package.GetRelationshipsByType(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").FirstOrDefault();
                if (relationship == null)
                    throw new ArgumentException("Not an Open XML Document (required relationship does not exist).");

                PackagePart part = package.GetPart(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri));
                switch (part.ContentType)
                {
                    case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
                        DocumentType = typeof(WordprocessingDocument);
                        break;
                    case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                        DocumentType = typeof(SpreadsheetDocument);
                        break;
                    case "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":
                    case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
                        DocumentType = typeof(PresentationDocument);
                        break;
                    default:
                        throw new ArgumentException("Not an Open XML Document (unsupported content type: " + part.ContentType + ").");
                }
            }
        }

        #endregion Constructors and Initializers

        #region Package-related methods

        /// <summary>
        /// Opens the <see cref="Package"/> on this stream.
        /// </summary>
        /// <returns></returns>
        public Package OpenPackage()
        {
            return Package.Open(this);
        }

        /// <summary>
        /// Opens the <see cref="Package"/> on this stream, using a given file mode.
        /// </summary>
        /// <param name="packageMode"></param>
        /// <returns></returns>
        public Package OpenPackage(FileMode packageMode)
        {
            return Package.Open(this, packageMode);
        }

        /// <summary>
        /// Opens the <see cref="Package"/> on this stream, using a given file mode
        /// and file access setting.
        /// </summary>
        /// <param name="packageMode"></param>
        /// <param name="packageAccess"></param>
        /// <returns></returns>
        public Package OpenPackage(FileMode packageMode, FileAccess packageAccess)
        {
            return Package.Open(this, packageMode, packageAccess);
        }

        #endregion Package-related methods

        #region OpenXmlPackage-related methods

        /// <summary>
        /// Gets the type of the <see cref="OpenXmlPackage"/> stored on this 
        /// stream, i.e., either <see cref="WordprocessingDocument"/>,
        /// <see cref="SpreadsheetDocument"/>, <see cref="PresentationDocument"/>-
        /// </summary>
        public Type DocumentType { get; private set; }

        /// <summary>
        ///  Creates a new instance of a base class of the <see cref="OpenXmlPackage"/>
        ///  class from this stream. 
        /// </summary>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public OpenXmlPackage OpenDocument(bool isEditable)
        {
            return OpenDocument(isEditable, new OpenSettings());
        }

        /// <summary>
        ///  Creates a new instance of a base class of the <see cref="OpenXmlPackage"/>
        ///  class from this stream. 
        /// </summary>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public abstract OpenXmlPackage OpenDocument(bool isEditable, OpenSettings openSettings);

        #endregion OpenXmlPackage-related methods

        #region File-related properties and methods

        /// <summary>
        /// Gets the path of the file from which the stream was initialized or 
        /// to which it was last saved.
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Saves the contents of the stream to the file from which it was initialized.
        /// If the stream was not initialized from a file, use <see cref="SaveAs"/> instead.
        /// </summary>
        /// <exception cref="ArgumentException">If the stream was not initialized from a file</exception>
        public void Save()
        {
            if (Path == null)
                throw new ArgumentException(
                    "Path is not defined (stream not initialized from file or never saved using SaveAs).");

            SaveAs(Path);
        }

        /// <summary>
        /// Saves the contents of the stream to the given file. 
        /// The given path will be used in subsequent invocations of the 
        /// <see cref="Save"/> method.
        /// </summary>
        /// <param name="path">The file path</param>
        public void SaveAs(string path)
        {
            if (path == null)
                throw new ArgumentNullException("path");

            // Write the entire contents of this MemoryStream to a FileStream.
            using (FileStream stream = new FileStream(path, FileMode.Create))
                WriteTo(stream);

            // Remember the path of the file to which we've last saved the stream.
            // The Path property will, therefore, always contain the path of the
            // file from which this stream was initialized or to which it was last
            // saved, i.e., the file which most closely resembles its contents.
            Path = path;
        }

        #endregion File-related properties and methods
    }
}
