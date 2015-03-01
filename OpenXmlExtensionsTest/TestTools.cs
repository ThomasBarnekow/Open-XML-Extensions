/*
 * TestTools.cs - NUnit Test Tools
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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Packaging;

using NUnit.Framework;

namespace OpenXmlExtensionsTest
{
    public static class TestTools
    {
        public static void RemoveFiles(string path, string searchPattern)
        {
            DirectoryInfo directory = new DirectoryInfo(path);
            foreach (FileInfo file in directory.GetFiles(searchPattern))
                file.Delete();
        }

        public static void PrepareWordprocessingDocument(string path)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                foreach (OpenXmlPart part in document.GetAllParts())
                    if (part.RootElement != null)
                        part.RootElement.Save();
            }
        }

        public static void PrepareSpreadsheetDocument(string path)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                foreach (OpenXmlPart part in document.GetAllParts())
                    if (part.RootElement != null)
                        part.RootElement.Save();
            }
        }

        public static void PreparePresentationDocument(string path)
        {
            using (PresentationDocument document = PresentationDocument.Open(path, true))
            {
                foreach (OpenXmlPart part in document.GetAllParts())
                    if (part.RootElement != null)
                        part.RootElement.Save();
            }
        }

        /// <summary>
        /// Asserts that two OpenXmlPackage instances have the same content. 
        /// </summary>
        /// <param name="first">The first OpenXmlPackage</param>
        /// <param name="second">The second OpenXmlPackage</param>
        public static void AssertThatPackagesAreEqual(OpenXmlPackage first, OpenXmlPackage second)
        {
            List<OpenXmlPart> firstParts = first.GetAllParts().ToList();
            List<OpenXmlPart> secondParts = second.GetAllParts().ToList();

            // Assert that we have an equivalent list of parts.
            Assert.That(
                firstParts.Select(p => p.GetType()),
                Is.EquivalentTo(
                    secondParts.Select(p => p.GetType())));

            // Assert that the parts' root elements are equal.
            for (int i = 0; i < firstParts.Count(); i++)
            {
                if (firstParts[i].ContentType.EndsWith("xml"))
                {
                    string firstString = GetXmlString(firstParts[i]);   // firstParts[i].GetRootElement().ToString();
                    string secondString = GetXmlString(secondParts[i]);   // secondParts[i].GetRootElement().ToString();

                    Assert.That(firstString, Is.EqualTo(secondString));
                }
                else
                {
                    byte[] firstByteArray;
                    using (Stream stream = firstParts[i].GetStream())
                    using (BinaryReader binaryReader = new BinaryReader(stream))
                    {
                        int len = (int)binaryReader.BaseStream.Length;
                        firstByteArray = binaryReader.ReadBytes(len);
                    }

                    byte[] secondByteArray;
                    using (Stream stream = secondParts[i].GetStream())
                    using (BinaryReader binaryReader = new BinaryReader(stream))
                    {
                        int len = (int)binaryReader.BaseStream.Length;
                        secondByteArray = binaryReader.ReadBytes(len);
                    }

                    Assert.That(firstByteArray, Is.EquivalentTo(secondByteArray));
                }
            }
        }

        /// <summary>
        /// Lets the given part's RootElement produce the XML string.
        /// </summary>
        /// <param name="part"></param>
        /// <returns></returns>
        public static string GetXmlString(OpenXmlPart part)
        {
            StringBuilder sb = new StringBuilder();
            using (XmlWriter xw = XmlWriter.Create(sb))
            {
                if (part.RootElement != null)
                    part.RootElement.WriteTo(xw);
                else
                    sb.Append(string.Empty);
            }
            return sb.ToString();
        }
    }
}
