/*
 * FlatOpcPackagePart.cs - PackagePart for FlatOpcPackage
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

using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
    /// <summary>
    /// This class represents a <see cref="PackagePart"/> contained in a 
    /// <see cref="FlatOpcPackage"/>.
    /// </summary>
    public class FlatOpcPackagePart : PackagePart
    {
        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        
        private XDeclaration _declaration = new XDeclaration("1.0", "UTF-8", "yes");

        private FlatOpcPackage _package = null;

        private XDocument _partDocument = null;
        byte[] _partBinaryData;

        Uri _partUri = null;
        string _partContentType = null;

        // private XElement _partElement = null;

        // This is one of the base class constructors. However, we actually don't 
        // support it here, because we always need a contentType.
        // internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri)
        //     : this(package, partUri, null, CompressionOption.NotCompressed)
        // { }

        /// <summary>
        /// Initializes a new instance of FlatOpcPackagePart with a URI, a content type,
        /// and <see cref="CompressionOption.NotCompressed"/>.
        /// </summary>
        /// <param name="package">The container <see cref="FlatOpcPackage"/>.</param>
        /// <param name="partUri">This part's <see cref="Uri"/>.</param>
        /// <param name="contentType">This part's content type.</param>
        internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri, string contentType)
            : this(package, partUri, contentType, CompressionOption.NotCompressed)
        { }

        /// <summary>
        /// Initializes a new instance of FlatOpcPackagePart with a URI, a content type,
        /// and a compression option.
        /// </summary>
        /// <remarks>
        /// For Flat OPC documents, we only support <see cref="CompressionOption.NotCompressed"/>,
        /// so all other compression options are simply ignored. 
        /// </remarks>
        /// <param name="package">The container <see cref="FlatOpcPackage"/>.</param>
        /// <param name="partUri">This part's <see cref="Uri"/>.</param>
        /// <param name="contentType">This part's content type.</param>
        /// <param name="compressionOption">The compression option.</param>
        internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri, string contentType, CompressionOption compressionOption)
            : base(package, partUri, contentType, compressionOption)
        {
            if (package == null)
                throw new ArgumentNullException("package");
            if (partUri == null)
                throw new ArgumentNullException("partUri");
            if (contentType == null)
                throw new ArgumentNullException("contentType");

            // Let's just ignore compression options. OpenXml packages use them,
            // so we'd run into errors if we checked them here.
            // if (compressionOption != CompressionOption.NotCompressed)
            //     throw new NotSupportedException("CompressionOption is not supported: " + compressionOption);

            _package = package;

            // We remember these so we can use them while the part gets disposed.
            _partUri = partUri;
            _partContentType = contentType;
        }
        
        /// <summary>
        /// Returns the underlying stream that is represented by this part 
        /// in the specified mode with the access.
        /// </summary>
        /// <param name="mode">The file mode.</param>
        /// <param name="access">The file access.</param>
        /// <returns></returns>
        protected override Stream GetStreamCore(FileMode mode, FileAccess access)
        {
            FlatOpcPackagePartStream stream = null;
            if (mode == FileMode.Open || mode == FileMode.OpenOrCreate)
            {
                // Create stream with ReadWrite access regardless of the type of
                // access requested. This is required to write the existing part
                // to the stream. Write PartDocument to stream if it is not empty.
                stream = new FlatOpcPackagePartStream(this, FileAccess.ReadWrite);
                if (ContentType.EndsWith("xml"))
                {
                    if (PartDocument != null && PartDocument.Root != null)
                    {
                        PartDocument.Save(stream);
                        stream.Position = 0;
                    }
                    else if (mode == FileMode.Open)
                    {
                        throw new IOException("Part is empty.");
                    }
                }
                else
                {
                    if (PartBinaryData != null)
                    {
                        stream.Write(PartBinaryData, 0, PartBinaryData.Length);
                        stream.Position = 0;
                    }
                    else if (mode == FileMode.Open)
                    {
                        throw new IOException("Part is empty.");
                    }
                }

                // Set the desired access level, i.e., possibly reducing it to
                // Read only.
                stream.Access = access;
            }
            else if (mode == FileMode.Create)
            {
                if (ContentType.EndsWith("xml"))
                    PartDocument = null;
                else
                    PartBinaryData = null;

                stream = new FlatOpcPackagePartStream(this, access);
            }
            else if (mode == FileMode.CreateNew)
            {
                if (ContentType.EndsWith("xml"))
                    if (PartDocument != null)
                        throw new IOException("XML part is not empty.");
                else 
                    if (PartBinaryData != null)
                        throw new IOException("Binary part is not empty.");

                stream = new FlatOpcPackagePartStream(this, access);
            }
            else
            {
                throw new IOException("Unsupported FileMode: " + mode);
            }
            return stream;
        }

        /// <summary>
        /// Gets or sets the root <see cref="XElement"/> of the <see cref="XDocument"/>
        /// contained in this part. This is used by FlatOpcPackage to initialize this
        /// FlatOpcPackagePart in case the content type is XML.
        /// </summary>
        internal XElement RootElement
        {
            get
            {
                if (_partDocument != null)
                    return _partDocument.Root;
                else
                    return null;
            }

            set
            {
                if (!ContentType.EndsWith("xml"))
                    throw new ArgumentException("Can't set the RootElement if content type is '" + ContentType + "'");

                PartDocument = new XDocument(_declaration, value);
            }
        }

        /// <summary>
        /// Get's or sets the <see cref="XDocument"/> contained in this part.
        /// </summary>
        internal XDocument PartDocument 
        {
            get
            {
                return _partDocument;
            }

            set
            {
                if (!ContentType.EndsWith("xml"))
                    throw new ArgumentException("Can't set PartDocument if content type is '" + ContentType + "'");

                _partDocument = value;

                // This is the eager initialization variant.
                //_partElement = new XElement(pkg + "part",
                //    new XAttribute(pkg + "name", Uri),
                //    new XAttribute(pkg + "contentType", ContentType),
                //    new XElement(pkg + "xmlData",
                //        RootElement));
            }
        }

        internal byte[] PartBinaryData
        {
            get
            {
                return _partBinaryData;
            }

            set
            {
                if (ContentType.EndsWith("xml"))
                    throw new ArgumentException("Can't set PartDocument if content type is '" + ContentType + "'");

                _partBinaryData = value;

                // This is the eager initialization variant.
                //// The following expression creates the base64String, then chunks
                //// it to lines of 76 characters long.
                //string base64String = System.Convert.ToBase64String(_partBinaryData)
                //    .Select((c, i) => new { Character = c, Chunk = i / 76 })
                //    .GroupBy(c => c.Chunk)
                //    .Aggregate(
                //        new StringBuilder(),
                //        (s, i) =>
                //            s.Append(
                //                i.Aggregate(
                //                    new StringBuilder(),
                //                    (seed, it) => seed.Append(it.Character),
                //                    sb => sb.ToString())).Append(Environment.NewLine),
                //        s => s.ToString());

                //_partElement =  new XElement(pkg + "part",
                //    new XAttribute(pkg + "name", Uri),
                //    new XAttribute(pkg + "contentType", ContentType),
                //    new XAttribute(pkg + "compression", "store"),
                //    new XElement(pkg + "binaryData", base64String));
            }
        }

        /// <summary>
        /// Gets the <see cref="XElement"/> representing this part in a Flat OPC
        /// package. This is used by the <see cref="FlatOpcPackage"/> to assemble
        /// the Flat OPC document.
        /// </summary>
        internal XElement PartElement
        {
            get
            {
                // This is the lazy/late initialization variant.
                if (_partContentType.EndsWith("xml"))
                {
                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", _partUri),
                        new XAttribute(pkg + "contentType", _partContentType),
                        new XElement(pkg + "xmlData",
                            RootElement));
                }
                else
                {
                    // The following expression creates the base64String, then chunks
                    // it to lines of 76 characters long.
                    string base64String = System.Convert.ToBase64String(_partBinaryData)
                        .Select((c, i) => new { Character = c, Chunk = i / 76 })
                        .GroupBy(c => c.Chunk)
                        .Aggregate(
                            new StringBuilder(),
                            (s, i) =>
                                s.Append(
                                    i.Aggregate(
                                        new StringBuilder(),
                                        (seed, it) => seed.Append(it.Character),
                                        sb => sb.ToString())).Append(Environment.NewLine),
                            s => s.ToString());

                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", _partUri),
                        new XAttribute(pkg + "contentType", _partContentType),
                        new XAttribute(pkg + "compression", "store"),
                        new XElement(pkg + "binaryData", base64String));
                }

                // This is the old code used when we did eager initialization.
                // return _partElement;
            }
        }
    }
}
