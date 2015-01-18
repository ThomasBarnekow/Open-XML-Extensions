using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transforms
{
    /// <summary>
    /// Utility class providing a number of tools for transforms.
    /// </summary>
    public static class TransformTools
    {
        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC string.
        /// </summary>
        /// <param name="text">The Flat OPC string.</param>
        /// <typeparam name="DocumentType">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static DocumentType FromFlatOpcString<DocumentType>(string text)
            where DocumentType : OpenXmlPackage
        {
            Type t = typeof(DocumentType);
            if (t == typeof(WordprocessingDocument))
                return WordprocessingDocument.FromFlatOpcString(text) as DocumentType;
            else if (t == typeof(SpreadsheetDocument))
                return SpreadsheetDocument.FromFlatOpcString(text) as DocumentType;
            else if (t == typeof(PresentationDocument))
                return PresentationDocument.FromFlatOpcString(text) as DocumentType;
            else
                throw new OpenXmlTransformException("Unsupported document type: " + t);
        }

        /// <summary>
        /// Creates a new instance of DocumentType from a Flat OPC <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The Flat OPC <see cref="XDocument"/>.</param>
        /// <typeparam name="DocumentType">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
        /// <returns>A new instance of DocumentType</returns>
        public static DocumentType FromFlatOpcDocument<DocumentType>(XDocument document)
            where DocumentType : OpenXmlPackage
        {
            Type t = typeof(DocumentType);
            if (t == typeof(WordprocessingDocument))
                return WordprocessingDocument.FromFlatOpcDocument(document) as DocumentType;
            else if (t == typeof(SpreadsheetDocument))
                return SpreadsheetDocument.FromFlatOpcDocument(document) as DocumentType;
            else if (t == typeof(PresentationDocument))
                return PresentationDocument.FromFlatOpcDocument(document) as DocumentType;
            else
                throw new OpenXmlTransformException("Unsupported document type: " + t);
        }
    }
}
