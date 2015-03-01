/*
 * FlatOpcPackagePartStream.cs - Stream for FlatOpcPackagePart
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
    /// Stream for <see cref="FlatOpcPackagePart"/>.
    /// </summary>
    internal class FlatOpcPackagePartStream : MemoryStream
    {
        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        FlatOpcPackagePart _part;
        bool _disposed = false;

        /// <summary>
        /// Initializes a new instance of FlatOpcPackagePartStream linked to the
        /// given <see cref="FlatOpcPackagePart"/> and having ReadWrite access.
        /// </summary>
        /// <param name="part">The <see cref="FlatOpcPackagePart"/> which created this stream.</param>
        internal FlatOpcPackagePartStream(FlatOpcPackagePart part)
            : this(part, FileAccess.ReadWrite)
        {
            _part = part;
        }

        /// <summary>
        /// Initializes a new instance of FlatOpcPackagePartStream linked to the
        /// given <see cref="FlatOpcPackagePart"/> and having the specificed 
        /// <see cref="FileAccess"/>.
        /// </summary>
        /// <param name="part">The <see cref="FlatOpcPackagePart"/> which created this stream.</param>
        /// <param name="access">Read, ReadWrite, or Write.</param>
        internal FlatOpcPackagePartStream(FlatOpcPackagePart part, FileAccess access)
            : base()
        {
            _part = part;
            Access = access;
        }

        /// <summary>
        /// Gets or sets the <see cref="FileAccess"/> mode.
        /// </summary>
        internal FileAccess Access { get; set; }

        /// <summary>
        /// Indicates whether the stream is readable.
        /// </summary>
        public override bool CanRead
        {
            get 
            { 
                return base.CanRead && 
                      (Access == FileAccess.Read || 
                       Access == FileAccess.ReadWrite); 
            }
        }

        /// <summary>
        /// Indicates whether the stream is writeable.
        /// </summary>
        public override bool CanWrite
        {
            get 
            {
                return base.CanWrite &&
                      (Access == FileAccess.ReadWrite ||
                       Access == FileAccess.Write);
            }
        }

        /// <summary>
        /// Replaces the <see cref="FlatOpcPackagePart.Document"/> with the 
        /// <see cref="XDocument"/> contained on this stream, unless we can't
        /// seek or read.
        /// </summary>
        public override void Flush()
        {
#if VERBOSE
            // This is for testing purposes only.
            Console.WriteLine("FlatOpcPackagePartStream.Flush(): " + _part.Uri);
#endif
            SaveToPart();
        }

        /// <summary>
        /// Disposes this stream, replacing the <see cref="FlatOpcPackagePart.Document"/>
        /// with the <see cref="XDocument"/> contained on this stream (unless the stream
        /// is already disposed or disposing is false or we can't seek or read).
        /// </summary>
        /// <param name="disposing">True when disposing, false otherwise.</param>
        protected override void Dispose(bool disposing)
        {
#if VERBOSE
             // This is for testing purposes only.
            Console.WriteLine("FlatOpcPackagePartStream.Dispose(" + disposing + "): " + _part.Uri);
#endif
            if (_disposed)
                return;

            if (disposing)
            {
                SaveToPart();
                _disposed = true;
            }

            base.Dispose(disposing);
        }


        /// <summary>
        /// Replaces the <see cref="FlatOpcPackagePart.Document"/> with the 
        /// <see cref="XDocument"/> contained on this stream, unless we can't
        /// seek or read.
        /// </summary>
        private void SaveToPart()
        {
            if (CanSeek && CanRead)
            {
                if (Length == 0)
                    return; 

                Position = 0;
                if (_part.ContentType.EndsWith("xml"))
                {
                    _part.PartDocument = XDocument.Load(this);
                }
                else
                {
                    byte[] buffer = new byte[Length];
                    Array.Copy(GetBuffer(), buffer, Length);
                    _part.PartBinaryData = buffer;

                    //The following code led to an endless loop because the BinaryReader
                    //disposed this stream while were saving to the part because somebody
                    //else disposed this stream. 
                    //using (BinaryReader binaryReader = new BinaryReader(this))
                    //{
                    //    int len = (int)binaryReader.BaseStream.Length;
                    //    _part.PartBinaryData = binaryReader.ReadBytes(len);
                    //}
                }
            }
        }
    }
}
