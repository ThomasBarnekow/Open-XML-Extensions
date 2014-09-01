using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.IO.Packaging;

using System.Xml;
using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
    internal class FlatOpcPackagePartStream : MemoryStream
    {
        FlatOpcPackagePart _part;
        bool _disposed = false;

        internal FlatOpcPackagePartStream(FlatOpcPackagePart part)
            : this(part, FileAccess.ReadWrite)
        {
            _part = part;
        }

        internal FlatOpcPackagePartStream(FlatOpcPackagePart part, FileAccess access)
            : base()
        {
            _part = part;
            Access = access;
        }

        internal FileAccess Access { get; set; }

        public override bool CanRead
        {
            get 
            { 
                return base.CanRead && 
                      (Access == FileAccess.Read || 
                       Access == FileAccess.ReadWrite); 
            }
        }

        public override bool CanWrite
        {
            get 
            {
                return base.CanWrite &&
                      (Access == FileAccess.ReadWrite ||
                       Access == FileAccess.Write);
            }
        }

        public override void Flush()
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackagePartStream: Flushing: " + _part.Uri);
#endif
            SaveDocument();
        }

        protected override void Dispose(bool disposing)
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackagePartStream: Dispose(" + disposing + "): " + _part.Uri);
#endif
            if (_disposed)
                return;

            if (disposing)
            {
                SaveDocument();
                _disposed = true;
            }

            base.Dispose(disposing);
        }

        private void SaveDocument()
        {
            Position = 0;
            if (Length > 0)
                _part.PartDocument = XDocument.Load(this);
        }
    }
}
