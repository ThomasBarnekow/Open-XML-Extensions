/*
 * XmlTransformations.cs - Transformations for XML and Open XML documents
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

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ContractArchitect.OpenXml.Transformation
{
    #region XML to Open XML transformations

    /// <summary>
    /// This interface defines methods implemented by transformations from generic XML to
    /// Open XML documents.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public interface IXmlToOpenXmlTransformation<out TDocument>
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Transforms a generic XML string to an Open XML document.
        /// </summary>
        /// <param name="text">The XML string.</param>
        /// <returns>The Open XML document.</returns>
        TDocument ToOpenXml(string text);

        /// <summary>
        /// Transforms a generic <see cref="XDocument" /> into an Open XML document.
        /// </summary>
        /// <param name="document">The <see cref="XDocument" />.</param>
        /// <returns>The Open XML document.</returns>
        TDocument ToOpenXml(XDocument document);
    }

    /// <summary>
    /// This class is the abstract base class for transformations from generic XML to Open XML
    /// documents that perform their transformation based on an XML <see cref="string" />.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class XmlStringToOpenXmlTransformation<TDocument> : IXmlToOpenXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        public abstract TDocument ToOpenXml(string text);

        public TDocument ToOpenXml(XDocument document)
        {
            return document == null ? null : ToOpenXml(document.ToString());
        }
    }

    /// <summary>
    /// This class is the abstract base class for transformations from generic XML to Open XML
    /// documents that perform their transformation based on an <see cref="XDocument" />.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class XDocumentToOpenXmlTransformation<TDocument> : IXmlToOpenXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        public TDocument ToOpenXml(string text)
        {
            return text == null ? null : ToOpenXml(XDocument.Parse(text));
        }

        public abstract TDocument ToOpenXml(XDocument document);
    }

    #endregion

    #region Open XML to XML transformations

    /// <summary>
    /// This interface defines methods implemented by transformations from Open XML to generic
    /// XML documents. All methods will perform the exact same transformation and only take
    /// the input in different formats.
    /// </summary>
    /// <seealso cref="FlatOpcStringToXmlTransformation{TDocument}" />
    /// <seealso cref="FlatOpcDocumentToXmlTransformation{TDocument}" />
    /// <seealso cref="OpenXmlPackageToXmlTransformation{TDocument}" />
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public interface IOpenXmlToXmlTransformation<in TDocument>
        where TDocument : OpenXmlPackage
    {
        XDocument ToXml(string text);
        XDocument ToXml(XDocument document);
        XDocument ToXml(TDocument packageDocument);
    }

    /// <summary>
    /// This class is the abstract base class for transformations from Open XML to generic
    /// XML documents that perform their transformation on the Flat OPC <see cref="string" />
    /// representation of an Open XML package.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcStringToXmlTransformation<TDocument> : IOpenXmlToXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        public abstract XDocument ToXml(string text);

        public XDocument ToXml(XDocument document)
        {
            return document == null ? null : ToXml(document.ToString());
        }

        public XDocument ToXml(TDocument packageDocument)
        {
            return packageDocument == null ? null : ToXml(packageDocument.ToFlatOpcString());
        }
    }

    /// <summary>
    /// This class is the abstract base class for transformations from Open XML to generic
    /// XML documents that perform their transformation on the Flat OPC <see cref="XDocument" />
    /// representation of an Open XML package, using the Linq to XML classes.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class FlatOpcDocumentToXmlTransformation<TDocument> : IOpenXmlToXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        public XDocument ToXml(string text)
        {
            return text == null ? null : ToXml(XDocument.Parse(text));
        }

        public abstract XDocument ToXml(XDocument document);

        public XDocument ToXml(TDocument packageDocument)
        {
            return packageDocument == null ? null : ToXml(packageDocument.ToFlatOpcDocument());
        }
    }

    /// <summary>
    /// This class is the abstract base class for transformations from Open XML to generic
    /// XML documents that perform their transformation on one of the subclasses of
    /// <see cref="OpenXmlPackage" />, using the Open XML SDK.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage" />.</typeparam>
    public abstract class OpenXmlPackageToXmlTransformation<TDocument> : IOpenXmlToXmlTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        public XDocument ToXml(string text)
        {
            return text == null ? null : ToXml(TransformationTools.FromFlatOpcString<TDocument>(text));
        }

        public XDocument ToXml(XDocument document)
        {
            return document == null ? null : ToXml(TransformationTools.FromFlatOpcDocument<TDocument>(document));
        }

        public abstract XDocument ToXml(TDocument packageDocument);
    }

    #endregion
}
