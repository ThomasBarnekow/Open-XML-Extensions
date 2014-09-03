/*
 * FlatOpcPackage.cs - Package for Flat OPC documents
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

using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
    /// <summary>
    /// This class represents a <see cref="Package"/> for Flat OPC documents.
    /// </summary>
    public class FlatOpcPackage : Package
    {
        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        // Default values for the Package.Open method overloads
        private static readonly FileMode _defaultFileMode = FileMode.OpenOrCreate;
        private static readonly FileAccess _defaultFileAccess = FileAccess.ReadWrite;
        private static readonly FileShare _defaultFileShare = FileShare.None;

        private static readonly FileMode _defaultStreamMode = FileMode.Open;
        private static readonly FileAccess _defaultStreamAccess = FileAccess.Read;

        private XDeclaration _declaration = new XDeclaration("1.0", "UTF-8", "yes");
        private XProcessingInstruction _processingInstruction = null;

        private SortedList<Uri, FlatOpcPackagePart> _partList = 
            new SortedList<Uri, FlatOpcPackagePart>(new UriComparer());

        private Stream _stream = null;
        private bool _disposed = false;

        /// <summary>
        /// Initializes a new instance of FlatOpcPackage with the given file access
        /// and non-streaming mode (i.e., streaming is false).
        /// </summary>
        /// <param name="openFileAccess">The desired <see cref="FileAccess"/> mode.</param>
        internal FlatOpcPackage(FileAccess openFileAccess)
            : this(openFileAccess, false)
        { }

        /// <summary>
        /// Initializes a new instance of FlatOpcPackage with the given file access
        /// and streaming mode.
        /// </summary>
        /// <remarks>
        /// Streaming is currently not supported. Therefore, if streaming is true,
        /// an <see cref="IOException"/> is thrown.
        /// </remarks>
        /// <exception cref="IOException">If streaming is true.</exception>
        /// <param name="openFileAccess"></param>
        /// <param name="streaming"></param>
        internal FlatOpcPackage(FileAccess openFileAccess, bool streaming)
            : base(openFileAccess, streaming)
        {
            if (streaming)
                throw new IOException("Streaming is currently not supported");
        }

        /// <summary>
        /// Opens a package from an <see cref="XDocument"/>. 
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/></param>
        /// <returns></returns>
        public static FlatOpcPackage Open(XDocument document)
        {
            FlatOpcPackage package = new FlatOpcPackage(FileAccess.ReadWrite, false);
            package.Document = document;
            return package;
        }

        /// <summary>
        /// Opens a FlatOpcPackage at the specified path. This method calls the overload 
        /// which accepts all the parameters with the following defaults:
        /// FileMode   - FileMode.OpenOrCreate
        /// FileAccess - FileAccess.ReadWrite
        /// FileShare  - FileShare.None
        /// </summary>
        /// <param name="path">Path to the package.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(string path)
        {
            return Open(path, _defaultFileMode, _defaultFileAccess);
        }

        /// <summary>
        /// Opens a FlatOpcPackage at the specified path in the given mode. This method 
        /// calls the overload which accepts all the parameters with the following 
        /// defaults:
        /// FileAccess - FileAccess.ReadWrite
        /// FileShare  - FileShare.None
        /// </summary>
        /// <param name="path">Path to the package.</param>
        /// <param name="packageMode">FileMode in which the package should be opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(string path, FileMode packageMode)
        {
            return Open(path, packageMode, _defaultFileAccess);
        }

        /// <summary>
        /// Opens a FlatOpcPackage at the specified path in the given mode with the 
        /// specified access. This method calls the overload which accepts all 
        /// the parameters with the following defaults:
        /// FileShare  - FileShare.None        
        /// </summary>
        /// <param name="path">Path to the package.</param>
        /// <param name="packageMode">FileMode in which the package should be opened.</param>
        /// <param name="packageAccess">FileMode in which the package should be opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(string path, FileMode packageMode, FileAccess packageAccess)
        {
            return Open(path, packageMode, packageAccess, _defaultFileShare);
        }

        /// <summary>
        /// Opens a FlatOpcPackage with the specified parameters.
        /// </summary>
        /// <param name="path">Path to the package.</param>
        /// <param name="packageMode">FileMode in which the package should be opened.</param>
        /// <param name="packageAccess">FileMode in which the package should be opened.</param>
        /// <param name="packageShare">FileShare with which the package is opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(string path, FileMode packageMode, FileAccess packageAccess, FileShare packageShare)
        {
            return Open(new FileStream(path, packageMode, packageAccess, packageShare), packageMode, packageAccess);
        }

        /// <summary>
        /// Opens a FlatOpcPackage on the given stream. This method calls the overload 
        /// which accepts all the parameters with the following defaults:
        /// FileMode   - FileMode.Open
        /// FileAccess - FileAccess.Read
        /// </summary>
        /// <param name="stream">Stream on which the package is to be opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(Stream stream)
        {
            return Open(stream, _defaultStreamMode, _defaultStreamAccess);
        }

        /// <summary>
        /// Opens a FlatOpcPackage on the given stream. This method calls the overload 
        /// which accepts all the parameters with the following defaults:
        /// FileAccess - FileAccess.ReadWrite
        /// </summary>
        /// <param name="stream">Stream on which the package is to be opened.</param>
        /// <param name="packageMode">FileMode in which the package should be opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(Stream stream, FileMode packageMode)
        {
            return Open(stream, packageMode, _defaultStreamAccess);
        }

        /// <summary>
        /// Opens a FlatOpcPackage on the given stream. The package is opened in the 
        /// specified mode and with the access specified.
        /// </summary>
        /// <param name="stream">Stream on which the package is to be opened.</param>
        /// <param name="packageMode">FileMode in which the package is to be opened.</param>
        /// <param name="packageAccess">FileAccess on the package that is opened.</param>
        /// <returns>A new instance of FlatOpcPackage.</returns>
        public static new FlatOpcPackage Open(Stream stream, FileMode packageMode, FileAccess packageAccess)
        {
            FlatOpcPackage package = new FlatOpcPackage(packageAccess);
            package.Init(stream, packageMode);
            return package;
        }

        /// <summary>
        /// Initializes this FlatOpcPackage.
        /// </summary>
        /// <param name="stream">The underlying <see cref="Stream"/>.</param>
        /// <param name="packageMode">The package's <see cref="FileMode"/>.</param>
        private void Init(Stream stream, FileMode packageMode)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            _stream = stream;

            if (packageMode == FileMode.Open || packageMode == FileMode.OpenOrCreate)
            {
                if (_stream.Length > 0)
                    LoadDocument();
                else if (packageMode == FileMode.OpenOrCreate)
                    SaveDocument();
                else
                    throw new IOException("Stream is empty");
            }
            else if (packageMode == FileMode.Create)
            {
                SaveDocument();
            }
            else if (packageMode == FileMode.CreateNew)
            {
                if (_stream.Length > 0)
                    throw new IOException("Stream is not empty");

                SaveDocument();
            }
            else
            {
                throw new IOException("Unsupported FileMode: " + packageMode);
            }
        }

        /// <summary>
        /// Loads a Flat OPC <see cref="XDocument"/> from the underlying 
        /// <see cref="Stream"/>.
        /// </summary>
        private void LoadDocument()
        {
            // Don't do anything if the package isn't backed by a stream.
            if (_stream == null)
                return;

            if (_stream.CanSeek && _stream.CanRead)
            {
                _stream.Position = 0;
                Document = XDocument.Load(_stream);
            }
        }

        /// <summary>
        /// Saves the Flat OPC <see cref="XDocument"/> represented by this
        /// FlatOpcPackage to the underlying <see cref="Stream"/>, unless
        /// we can't seek or write.
        /// </summary>
        private void SaveDocument()
        {
            // Don't do anything if the package isn't backed by a stream.
            if (_stream == null)
                return;

            if (_stream.CanSeek && _stream.CanWrite)
            {
                _stream.Position = 0;
                Document.Save(_stream);
                _stream.Position = 0;
            }
        }

        /// <summary>
        /// Gets the Flat OPC <see cref="XDocument"/> represented by this Package.
        /// </summary>
        public XDocument Document
        {
            get
            {
                return new XDocument(
                    _declaration,
                    _processingInstruction,
                    new XElement(pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        GetPartsCore().Select(pp => ((FlatOpcPackagePart)pp).PartElement)));
            }

            internal set
            {
                if (value == null)
                    throw new ArgumentNullException("Document");
                if (value.Root.Name != pkg + "package")
                    throw new ArgumentException("Not a Flat OPC document", "Document");

                _processingInstruction = value.Nodes()
                    .Where(n => n.NodeType == XmlNodeType.ProcessingInstruction)
                    .FirstOrDefault() as XProcessingInstruction;

                foreach (XElement element in value.Root.Elements().Where(e => e.Name == pkg + "part"))
                {
                    Uri partUri = PackUriHelper.CreatePartUri(new Uri(element.Attribute(pkg + "name").Value, UriKind.Relative));
                    string contentType = element.Attribute(pkg + "contentType").Value;

                    FlatOpcPackagePart packagePart = new FlatOpcPackagePart(this, partUri, contentType);
                    if (contentType.EndsWith("xml"))
                    {
                        packagePart.RootElement = (XElement)element.Element(pkg + "xmlData").FirstNode;
                    }
                    else
                    {
                        string base64StringInChunks = (string)element.Element(pkg + "binaryData");
                        char[] base64CharArray = base64StringInChunks
                            .Where(c => c != '\r' && c != '\n').ToArray();
                        byte[] byteArray =
                            System.Convert.FromBase64CharArray(
                                base64CharArray, 0, base64CharArray.Length);
                        packagePart.PartBinaryData = byteArray;
                    }                       

                    _partList.Add(partUri, packagePart);
                }
            }
        }

        /// <summary>
        /// Creates a new <see cref="FlatOpcPackagePart"/> with the given part 
        /// <see cref="Uri"/>, content type, and <see cref="CompressionOption"/>.
        /// </summary>
        /// <param name="partUri">The <see cref="FlatOpcPackagePart"/>'s <see cref="Uri"/></param>
        /// <param name="contentType">The content type.</param>
        /// <param name="compressionOption">The <see cref="CompressionOption"/>.</param>
        /// <returns>A new instance of <see cref="FlatOpcPackagePart"/>.</returns>
        protected override PackagePart CreatePartCore(Uri partUri, string contentType, 
            CompressionOption compressionOption)
        {
            if (partUri == null)
                throw new ArgumentNullException("partUri");

            FlatOpcPackagePart packagePart = new FlatOpcPackagePart(this, partUri, contentType, compressionOption);
            // packagePart.RootElement = null;

            _partList.Add(partUri, packagePart);
            return packagePart;
        }

        /// <summary>
        /// Deletes the <see cref="FlatOpcPackagePart"/> identified by the given
        /// part <see cref="Uri"/>.
        /// </summary>
        /// <param name="partUri">The <see cref="FlatOpcPackagePart"/>'s <see cref="Uri"/>.</param>
        protected override void DeletePartCore(Uri partUri)
        {
            // Remove the part from the list.
            // QUESTION: Should we also call FlushCore() or SaveDocument()?
            _partList.Remove(partUri);
        }

        /// <summary>
        /// Saves the Flat OPC document to the underlying stream or file.
        /// </summary>
        protected override void FlushCore()
        {
#if VERBOSE
            // This is for testing purposes only.
            Console.WriteLine("FlatOpcPackage.FlushCore()");
#endif
            SaveDocument();
        }

        /// <summary>
        /// Gets the <see cref="FlatOpcPackagePart"/> with the given <see cref="Uri"/>
        /// or null if it does not exist.
        /// </summary>
        /// <param name="partUri">The <see cref="FlatOpcPackagePart"/>'s <see cref="Uri"/>.</param>
        /// <returns>The <see cref="FlatOpcPackagePart"/> or null.</returns>
        protected override PackagePart GetPartCore(Uri partUri)
        {
            if (_partList.ContainsKey(partUri))
                return _partList[partUri];
            else
                return null;
        }

        /// <summary>
        /// Produces an array of the <see cref="FlatOpcPackagePart"/>s contained in this
        /// package, sorted in ascending order by the <see cref="FlatOpcPackagePart"/>s'
        /// <see cref="Uri"/>s.
        /// </summary>
        /// <returns>The sorted array of package parts.</returns>
        protected override PackagePart[] GetPartsCore()
        {
            List<PackagePart> parts = new List<PackagePart>(_partList.Keys.Count);
            foreach (Uri partUri in _partList.Keys)
                parts.Add(_partList[partUri]);

            return parts.ToArray();
        }

        /// <summary>
        /// Disposes this package, saving the Flat OPC document to the underlying
        /// stream or file (unless the package has been disposed already).
        /// </summary>
        /// <param name="disposing">True when disposing, false otherwise.</param>
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
                return;
#if VERBOSE
            // This is for testing purposes only.
            Console.WriteLine("FlatOpcPackage.Dispose(" + disposing + ")");
#endif
            try
            {
                if (disposing)
                {                   
                    SaveDocument();
                    if (_stream != null)
                        _stream.Dispose();

                    _stream = null;
                }
            }
            finally
            {
                _disposed = true;
                base.Dispose(disposing);
            }
        }
    }

    /// <summary>
    /// This class represents a <see cref="Uri"/> <see cref="Comparer"/>.
    /// </summary>
    internal class UriComparer : Comparer<Uri>
    {
        /// <summary>
        /// Compares two <see cref="Uri"/>s.
        /// </summary>
        /// <param name="x">First <see cref="Uri"/>.</param>
        /// <param name="y">Second <see cref="Uri"/>.</param>
        /// <returns></returns>
        public override int Compare(Uri x, Uri y)
        {
            if (x != null && y != null)
                return x.ToString().CompareTo(y.ToString());
            else if (x == null && y == null)
                return 0;
            else
                throw new ArgumentNullException();
        }
    }
}
