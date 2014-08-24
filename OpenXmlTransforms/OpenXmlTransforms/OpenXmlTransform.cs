/*
 * OpenXmlTransform.cs - Transform for Open XML Documents
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
using System.IO.Packaging;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Transforms.Extensions;
using OpenXmlPowerTools;

namespace DocumentFormat.OpenXml.Transforms
{
    public abstract class OpenXmlTransform
    {
        #region Tools

        public static string ToFlatOpcString(XDocument document)
        {
            if (document == null)
                return null;

            return document.ToString();
        }

        public static string ToFlatOpcString(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            return ToFlatOpcDocument(document).ToString();
        }

        public static XDocument ToFlatOpcDocument(string xml)
        {
            if (xml == null)
                return null;

            return XDocument.Parse(xml);
        }

        public static XDocument ToFlatOpcDocument(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            return FlatOpc.OpcToFlatOpc(document);
        }

        public static WordprocessingDocument ToWordprocessingDocument(string xml)
        {
            if (xml == null)
                return null;

            return ToWordprocessingDocument(XDocument.Parse(xml));
        }

        public static WordprocessingDocument ToWordprocessingDocument(XDocument document)
        {
            if (document == null)
                return null;

            // Write OPC document to memory stream.
            MemoryStream stream = new MemoryStream();
            using (Package package = Package.Open(stream, FileMode.Create))
                FlatOpc.FlatOpcToOpc(document, package);

            // Create editable WordprocessingDocument from memory stream.
            return WordprocessingDocument.Open(stream, true);
        }

        #endregion

        #region Transforms

        public virtual string Transform(string xml)
        {
            return xml;
        }

        public virtual XDocument Transform(XDocument document)
        {
            return document;
        }

        public virtual WordprocessingDocument Transform(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            return TransformInPlace(document.Copy());
        }

        public virtual WordprocessingDocument TransformInPlace(WordprocessingDocument document)
        {
            return document;
        }

        #endregion
    }

    public abstract class FlatOpcStringTransform : OpenXmlTransform
    {
        public FlatOpcStringTransform()
            : base()
        { }

        sealed public override XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            string result = Transform(ToFlatOpcString(document));
            return ToFlatOpcDocument(result);
        }

        sealed public override WordprocessingDocument TransformInPlace(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            string result = Transform(ToFlatOpcString(document));
            using (WordprocessingDocument resultDoc = ToWordprocessingDocument(result))
                return document.ReplaceWith(resultDoc);
        }
    }

    public abstract class FlatOpcDocumentTransform : OpenXmlTransform
    {
        public FlatOpcDocumentTransform()
            : base()
        { }

        sealed public override string Transform(string xml)
        {
            if (xml == null)
                return null;

            XDocument result = Transform(ToFlatOpcDocument(xml));
            return ToFlatOpcString(result);
        }

        sealed public override WordprocessingDocument TransformInPlace(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            XDocument result = Transform(ToFlatOpcDocument(document));
            using (WordprocessingDocument wordDoc = ToWordprocessingDocument(result))
                return document.ReplaceWith(wordDoc);
        }
    }

    public abstract class WordprocessingDocumentTransform : OpenXmlTransform
    {
        public WordprocessingDocumentTransform()
            : base()
        { }

        public WordprocessingDocumentTransform(WordprocessingDocument template)
            : this()
        {
            Template = template;
        }

        public virtual WordprocessingDocument Template { get; set; }

        sealed public override string Transform(string xml)
        {
            if (xml == null)
                return null;

            using (WordprocessingDocument document = ToWordprocessingDocument(xml))
                return ToFlatOpcString(TransformInPlace(document));
        }

        sealed public override XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            using (WordprocessingDocument wordDoc = ToWordprocessingDocument(document))
                return ToFlatOpcDocument(TransformInPlace(wordDoc));
        }
    }
}
