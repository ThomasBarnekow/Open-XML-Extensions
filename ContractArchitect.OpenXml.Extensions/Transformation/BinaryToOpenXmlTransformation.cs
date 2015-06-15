/*
 * BinaryToOpenXmlTransformation.cs - Binary To OpenXML Transformation
 *
 * Copyright 2015 Thomas Barnekow
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
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ContractArchitect.OpenXml.Extensions;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.WordprocessingMLMapping;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using DocumentFormat.OpenXml.Wordprocessing;
using Packaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ContractArchitect.OpenXml.Transformation
{
    public static class BinaryToOpenXmlTransformation
    {
        public static readonly List<string> BinaryWordFileExtensions = new List<string> { ".doc", ".dot" };

        public static string Transform(string sourceFileName)
        {
            return Transform(sourceFileName, null);
        }

        public static string Transform(string sourceFileName, string destFileName)
        {
            if (sourceFileName == null)
                throw new ArgumentNullException("sourceFileName");
            
            if (!BinaryWordFileExtensions.Contains(Path.GetExtension(sourceFileName)))
                throw new ArgumentException(
                    "Incorrect extension for Microsoft Word 97-2003 file",
                    "sourceFileName");

            var sourceFileInfo = new FileInfo(sourceFileName);
            var sourcePath = sourceFileInfo.FullName;

            var destFileInfo = destFileName != null
                ? new FileInfo(destFileName)
                : new FileInfo(Path.ChangeExtension(sourcePath, ".docx"));
            var destPath = destFileInfo.FullName;

            // Initialize conforming destination path, assuming the one provided or determined
            // without having looked at the actual .doc file is the correct one.
            var conformingDestPath = destPath;            

            try
            {
                using (var reader = new StructuredStorageReader(sourcePath))
                {
                    // Parse input document.
                    var doc = new WordDocument(reader);

                    // Prepare output document.
                    var documentType = Converter.DetectOutputType(doc);
                    conformingDestPath = Converter.GetConformFilename(destPath, documentType);
                    var docx = DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(conformingDestPath, documentType);

                    // Convert.
                    Converter.Convert(doc, docx);

                    // Perform postprocessing.
                    using (var wordDocument = Packaging.WordprocessingDocument.Open(
                        conformingDestPath, true))
                    {
                        PostprocessWordprocessingDocument(wordDocument);
                    }

                    return conformingDestPath;
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (FileNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (ReadBytesAmountMismatchException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", sourceFileName);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (MagicNumberException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", sourceFileName);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (UnspportedFileVersionException ex)
            {
                TraceLogger.Error("File {0} has been created with a Word version older than Word 97.", sourceFileName);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (ByteParseException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", sourceFileName);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (MappingException ex)
            {
                TraceLogger.Error("There was an error while converting file {0}: {1}", sourceFileName, ex.Message);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (ZipCreationException ex)
            {
                TraceLogger.Error("Could not create output file {0}.", conformingDestPath);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed.", sourceFileName);
                TraceLogger.Debug(ex.ToString());
                throw;
            }
        }

        #region Postprocessing

        // These postprocessing methods are using Linq to XML rather than the SDK because previously we
        // had to process non-compliant Open XML markup produced by the Binary to Open XML Translator.
        // After having fixed the translator, the postprocessing could be streamlined.
        private static void PostprocessWordprocessingDocument(Packaging.WordprocessingDocument wordDocument)
        {
            if (wordDocument == null)
                return;

            var mainDocumentPart = wordDocument.MainDocumentPart;
            if (mainDocumentPart == null)
                return;

            var document = mainDocumentPart.GetRootElement();
            mainDocumentPart.SetRootElement((XElement) TransformDocument(document));
        }

        private static object TransformDocument(XNode node)
        {
            var element = node as XElement;
            if (element == null) return node;

            if (element.Name == W.r)
            {
                var children = element.Elements().ToList();
                if (children.Count == 1 && children[0].Name == W.t && children[0].IsEmpty)
                    return null;
            }

            return new XElement(element.Name, element.Attributes(),
                element.Nodes().Select(TransformDocument));
        }

         #endregion
    }
}
