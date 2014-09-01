using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
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

        internal FlatOpcPackage(FileAccess openFileAccess)
            : this(openFileAccess, false)
        { }

        internal FlatOpcPackage(FileAccess openFileAccess, bool streaming)
            : base(openFileAccess, streaming)
        {
            if (streaming)
                throw new IOException("Streaming is currently not supported");
        }

        public static new FlatOpcPackage Open(string path)
        {
            return Open(path, _defaultFileMode, _defaultFileAccess);
        }

        public static new FlatOpcPackage Open(string path, FileMode packageMode)
        {
            return Open(path, packageMode, _defaultFileAccess);
        }

        public static new FlatOpcPackage Open(string path, FileMode packageMode, FileAccess packageAccess)
        {
            return Open(path, packageMode, packageAccess, _defaultFileShare);
        }

        public static new FlatOpcPackage Open(string path, FileMode packageMode, FileAccess packageAccess, FileShare packageShare)
        {
            return Open(new FileStream(path, packageMode, packageAccess, packageShare), packageMode, packageAccess);
        }

        public static new FlatOpcPackage Open(Stream stream)
        {
            return Open(stream, _defaultStreamMode, _defaultStreamAccess);
        }

        public static new FlatOpcPackage Open(Stream stream, FileMode packageMode)
        {
            return Open(stream, packageMode, _defaultStreamAccess);
        }

        public static new FlatOpcPackage Open(Stream stream, FileMode packageMode, FileAccess packageAccess)
        {
            FlatOpcPackage package = new FlatOpcPackage(packageAccess);
            package.Init(stream, packageMode);
            return package;
        }

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

        private void LoadDocument()
        {
            if (_stream == null)
                throw new ArgumentNullException("stream");

            if (_stream.CanSeek && _stream.CanRead)
            {
                _stream.Position = 0;
                Document = XDocument.Load(_stream);
            }
        }

        private void SaveDocument()
        {
            if (_stream == null)
                throw new ArgumentNullException("stream");

            if (_stream.CanSeek && _stream.CanWrite)
            {
                _stream.Position = 0;
                Document.Save(_stream);
                _stream.Position = 0;
            }
        }

        public XDocument Document
        {
            get
            {
#if VERBOSE
                Console.WriteLine("FlatOpcPackage: Document");
#endif
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
                    packagePart.RootElement = (XElement)element.Element(pkg + "xmlData").FirstNode;

                    _partList.Add(partUri, packagePart);
                }
            }
        }

        protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption)
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: CreatePartCore: " + partUri);
#endif
            if (partUri == null)
                throw new ArgumentNullException("partUri");

            FlatOpcPackagePart packagePart = new FlatOpcPackagePart(this, partUri, contentType, compressionOption);
            packagePart.RootElement = null;

            _partList.Add(partUri, packagePart);
            return packagePart;
        }

        protected override void DeletePartCore(Uri partUri)
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: DeletePartCore: " + partUri);
#endif
            _partList.Remove(partUri);
        }

        protected override void FlushCore()
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: FlushCore");
#endif
            SaveDocument();
        }

        protected override PackagePart GetPartCore(Uri partUri)
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: GetPartCore: " + partUri);
#endif
            if (_partList.ContainsKey(partUri))
                return _partList[partUri];
            else
                return null;
        }

        protected override PackagePart[] GetPartsCore()
        {
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: GetPartsCore");
#endif
            List<PackagePart> parts = new List<PackagePart>(_partList.Keys.Count);
            foreach (Uri partUri in _partList.Keys)
                parts.Add(_partList[partUri]);

            return parts.ToArray();
        }

        protected override void Dispose(bool disposing)
        {
            if (_disposed)
                return;
#if VERBOSE
            Console.WriteLine("FlatOpcPackage: Dispose(" + disposing + ")");
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

    internal class UriComparer : Comparer<Uri>
    {
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
