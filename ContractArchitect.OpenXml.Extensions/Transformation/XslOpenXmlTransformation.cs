/*
 * XslOpenXmlTransformation.cs - XSL Open XML Transformation
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

using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using DocumentFormat.OpenXml.Packaging;

namespace ContractArchitect.OpenXml.Transformation
{
    /// <summary>
    /// This class uses XSL Stylesheets to perform Open XML transforms.
    /// </summary>
    /// <typeparam name="TDocument">A subclass of <see cref="OpenXmlPackage"/>.</typeparam>
    public class XslOpenXmlTransformation<TDocument> : FlatOpcStringTransformation<TDocument>
        where TDocument : OpenXmlPackage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XslOpenXmlTransformation{TDocument}" /> class.
        /// </summary>
        public XslOpenXmlTransformation()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XslOpenXmlTransformation{TDocument}" /> class
        /// with a <see cref="XslCompiledTransform" /> and an <see cref="XsltArgumentList" />.
        /// </summary>
        /// <param name="xslt">The compiled XSL transformation.</param>
        /// <param name="arguments">The arguments to be passed to the compiled XSL transformation.</param>
        public XslOpenXmlTransformation(XslCompiledTransform xslt, XsltArgumentList arguments)
        {
            Xslt = xslt;
            Arguments = arguments;
        }

        /// <summary>
        /// Gets or sets the compiled XSL transformation.
        /// </summary>
        public XslCompiledTransform Xslt { get; set; }

        /// <summary>
        /// Gets or sets the arguments to be passed to the compiled XSL transformation.
        /// </summary>
        public XsltArgumentList Arguments { get; set; }

        /// <summary>
        /// Transforms the XML string using the <see cref="XslCompiledTransform" />
        /// defined by the instance's <see cref="XslOpenXmlTransformation{TDocument}.Xslt" />
        /// property and the <see cref="XsltArgumentList" /> defined in the instance's
        /// <see cref="XslOpenXmlTransformation{TDocument}.Arguments" /> property.
        /// </summary>
        /// <param name="xml">The XML string to be transformed</param>
        /// <returns>The transformed XML string</returns>
        public override string Transform(string xml)
        {
            if (xml == null)
                return null;
            if (Xslt == null)
                throw new OpenXmlTransformationException("Xslt is not initialized");

            var sb = new StringBuilder();

            using (var sr = new StringReader(xml))
            using (var xr = XmlReader.Create(sr))
            using (var xw = XmlWriter.Create(sb))
                Xslt.Transform(xr, Arguments, xw);

            return sb.ToString();
        }
    }
}