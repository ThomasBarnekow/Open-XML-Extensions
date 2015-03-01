using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transforms
{
    /// <summary>
    /// Utility class providing a number of tools for transforms.
    /// </summary>
    public static class TransformTools
    {
        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC <see cref="XDocument" />.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument" />.</param>
        /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static TDocument FromFlatOpcDocument<TDocument>(XDocument document)
            where TDocument : OpenXmlPackage
        {
            var t = typeof (TDocument);
            if (t == typeof (WordprocessingDocument))
            {
                return WordprocessingDocument.FromFlatOpcDocument(document) as TDocument;
            }
            if (t == typeof (SpreadsheetDocument))
            {
                return SpreadsheetDocument.FromFlatOpcDocument(document) as TDocument;
            }
            if (t == typeof (PresentationDocument))
            {
                return PresentationDocument.FromFlatOpcDocument(document) as TDocument;
            }
            throw new OpenXmlTransformException("Unsupported document type: " + t);
        }

        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string.</param>
        /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static TDocument FromFlatOpcString<TDocument>(string text)
            where TDocument : OpenXmlPackage
        {
            var t = typeof (TDocument);
            if (t == typeof (WordprocessingDocument))
            {
                return WordprocessingDocument.FromFlatOpcString(text) as TDocument;
            }
            if (t == typeof (SpreadsheetDocument))
            {
                return SpreadsheetDocument.FromFlatOpcString(text) as TDocument;
            }
            if (t == typeof (PresentationDocument))
            {
                return PresentationDocument.FromFlatOpcString(text) as TDocument;
            }
            throw new OpenXmlTransformException("Unsupported document type: " + t);
        }
    }
}
