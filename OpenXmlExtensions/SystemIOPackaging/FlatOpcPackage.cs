using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
    public class FlatOpcPackage : Package
    {
        // Default values for the Package.Open method overloads
        private static readonly FileMode _defaultFileMode = FileMode.OpenOrCreate;
        private static readonly FileAccess _defaultFileAccess = FileAccess.ReadWrite;
        private static readonly FileShare _defaultFileShare = FileShare.None;

        private static readonly FileMode _defaultStreamMode = FileMode.Open;
        private static readonly FileAccess _defaultStreamAccess = FileAccess.Read;

        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        private XDeclaration _declaration = new XDeclaration("1.0", "UTF-8", "yes");
        private XProcessingInstruction _processingInstruction;

        private Dictionary<Uri, FlatOpcPackagePart> _packagePartDictionary = new Dictionary<Uri, FlatOpcPackagePart>();

        private Stream _stream;

        internal FlatOpcPackage(FileAccess openFileAccess)
            : base(openFileAccess)
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
            package._stream = stream;

            if (packageMode == FileMode.Open || packageMode == FileMode.OpenOrCreate)
            {
                if (stream.Length > 0)
                {
                    stream.Position = 0;
                    package.Document = XDocument.Load(stream);
                }
                else 
                {
                    if (packageMode == FileMode.OpenOrCreate)
                        package.SaveDocument();
                    else
                        throw new IOException("Stream is empty");
                }
            }
            else if (packageMode == FileMode.Create)
            {
                package.SaveDocument();
            }
            else if (packageMode == FileMode.CreateNew)
            {
                if (stream.Length > 0)
                    throw new IOException("Stream is not empty");

                package.SaveDocument();
            }
            else
            {
                throw new IOException("Unsupported FileMode: " + packageMode);
            }

            return package;
        }

        public XDocument Document
        {
            get
            {
                return new XDocument(
                    _declaration,
                    _processingInstruction,
                    new XElement(pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        _packagePartDictionary.Values.Select(pp => pp.PartElement)));
            }

            internal set
            {
                if (value == null)
                    throw new ArgumentNullException("Document");

                _processingInstruction = value.Nodes()
                    .Where(n => n.NodeType == XmlNodeType.ProcessingInstruction)
                    .FirstOrDefault() as XProcessingInstruction;

                foreach (XElement element in value.Root.Elements().Where(e => e.Name == pkg + "part"))
                {
                    Uri partUri = PackUriHelper.CreatePartUri(new Uri(element.Attribute(pkg + "name").Value, UriKind.Relative));
                    string contentType = element.Attribute(pkg + "contentType").Value;

                    FlatOpcPackagePart packagePart = new FlatOpcPackagePart(this, partUri, contentType);
                    packagePart.RootElement = (XElement)element.Element(pkg + "xmlData").FirstNode;

                    _packagePartDictionary.Add(partUri, packagePart);
                }
            }
        }

        protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption)
        {
            if (partUri == null)
                throw new ArgumentNullException("partUri");
            
            FlatOpcPackagePart packagePart = new FlatOpcPackagePart(this, partUri, contentType, compressionOption);
            packagePart.RootElement = null;

            _packagePartDictionary.Add(partUri, packagePart);
            return packagePart;
        }

        protected override void DeletePartCore(Uri partUri)
        {
            _packagePartDictionary.Remove(partUri);
        }

        protected override void FlushCore()
        {
            SaveDocument();
        }

        protected override PackagePart GetPartCore(Uri partUri)
        {
            if (_packagePartDictionary.ContainsKey(partUri))
                return _packagePartDictionary[partUri];
            else
                return null;
        }

        protected override PackagePart[] GetPartsCore()
        {
            return _packagePartDictionary.Values.ToArray();
        }

        private void SaveDocument()
        {
            if (_stream != null)
            {
                _stream.Position = 0;
                Document.Save(_stream);
                _stream.Position = 0;
            }
        }
    }
}
