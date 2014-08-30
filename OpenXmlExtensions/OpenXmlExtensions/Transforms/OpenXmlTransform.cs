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

using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Transforms
{
    public abstract class OpenXmlTransform
    {
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

            return TransformInPlace((WordprocessingDocument)document.Clone());
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

            string result = Transform(document.ToString());
            return XDocument.Parse(result);
        }

        sealed public override WordprocessingDocument TransformInPlace(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            string result = Transform(document.ToFlatOpcString());
            using (WordprocessingDocument resultDoc = WordprocessingDocument.FromFlatOpcString(result))
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

            XDocument result = Transform(XDocument.Parse(xml));
            return result.ToString();
        }

        sealed public override WordprocessingDocument TransformInPlace(WordprocessingDocument document)
        {
            if (document == null)
                return null;

            XDocument result = Transform(document.ToFlatOpcDocument());
            using (WordprocessingDocument wordDoc = WordprocessingDocument.FromFlatOpcDocument(result))
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

            using (WordprocessingDocument document = WordprocessingDocument.FromFlatOpcString(xml))
                return TransformInPlace(document).ToFlatOpcString();
        }

        sealed public override XDocument Transform(XDocument document)
        {
            if (document == null)
                return null;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.FromFlatOpcDocument(document))
                return TransformInPlace(wordDoc).ToFlatOpcDocument();
        }
    }
}
