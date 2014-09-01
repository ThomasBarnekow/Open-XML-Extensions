using System.Xml.Linq;

namespace System.IO.Packaging.FlatOpc
{
    public class FlatOpcPackagePart : PackagePart
    {
        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        
        private XDeclaration _declaration = new XDeclaration("1.0", "UTF-8", "yes");

        private FlatOpcPackage _package;

        internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri)
            : this(package, partUri, null, CompressionOption.NotCompressed)
        { }

        internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri, string contentType)
            : this(package, partUri, contentType, CompressionOption.NotCompressed)
        { }

        internal FlatOpcPackagePart(FlatOpcPackage package, Uri partUri, string contentType, CompressionOption compressionOption)
            : base(package, partUri, contentType, compressionOption)
        {
            // Let's just ignore compression options. OpenXml packages use them,
            // so we'd run into errors.
            // if (compressionOption != CompressionOption.NotCompressed)
            //     throw new NotSupportedException("CompressionOption is not supported: " + compressionOption);

            _package = package;
        }
        
        protected override Stream GetStreamCore(FileMode mode, FileAccess access)
        {
            FlatOpcPackagePartStream stream = null;
            if (mode == FileMode.Open || mode == FileMode.OpenOrCreate)
            {
                stream = new FlatOpcPackagePartStream(this);
                if (PartDocument != null)
                {
                    if (PartDocument.Root != null)
                    {
                        PartDocument.Save(stream);
                        stream.Position = 0;
                    }
                }
                else
                {
                    if (mode == FileMode.Open)
                        throw new IOException("PartDocument does not exist");
                }
                stream.Access = access;
            }
            else if (mode == FileMode.Create)
            {
                PartDocument = null;
                stream = new FlatOpcPackagePartStream(this, access);
            }
            else if (mode == FileMode.CreateNew)
            {
                if (PartDocument != null)
                    throw new IOException("PartDocument already exists");

                stream = new FlatOpcPackagePartStream(this, access);
            }
            else
            {
                throw new IOException("Unsupported FileMode: " + mode);
            }
            return stream;
        }

        internal XDocument PartDocument { get; set; }

        internal XElement RootElement
        {
            get
            {
                if (PartDocument != null)
                    return PartDocument.Root;
                else
                    return null;
            }

            set
            {
                PartDocument = new XDocument(_declaration, value);
            }
        }

        internal XElement PartElement
        {
            get
            {
                return new XElement(pkg + "part", 
                    new XAttribute(pkg + "name", Uri),
                    new XAttribute(pkg + "contentType", ContentType),
                    new XElement(pkg + "xmlData",
                        RootElement));
            }
        }
    }
}
