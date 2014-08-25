/*
 * PtMemoryStreamExtensions.cs - PowerTools for Open XML Extensions for MemoryStreams
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

using DocumentFormat.OpenXml.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Extensions for a number of Open XML SDK-related MemoryStreams, i.e.:
    /// (1) <see cref="OpenXmlMemoryStream"/>, 
    /// (2) <see cref="WordprocessingMemoryStream"/>, and
    /// (3) <see cref="SpreadsheetMemoryStream"/>.
    /// </summary>
    public static class PtMemoryStreamExtensions
    {
        #region OpenXmlMemoryStream-related methods

        /// <summary>
        /// Based on <see cref="OpenXmlMemoryStream.DocumentType"/>, this method 
        /// creates an instance of the corresponding child class of 
        /// <see cref="OpenXmlPowerToolsDocument"/> class from this stream.
        /// </summary>
        /// <returns></returns>
        public static OpenXmlPowerToolsDocument GetOpenXmlPowerToolsDocument(this OpenXmlMemoryStream stream)
        {
            if (stream.DocumentType == typeof(WordprocessingDocument))
                return new WmlDocument(stream.Path, stream);
            else if (stream.DocumentType == typeof(SpreadsheetDocument))
                return new SmlDocument(stream.Path, stream);
            else if (stream.DocumentType == typeof(PresentationDocument))
                return new PmlDocument(stream.Path, stream);
            else
                throw new ArgumentException("Not an Open XML Document: " + stream.DocumentType);
        }

        #endregion OpenXmlMemoryStream-related methods

        #region WordprocessingMemoryStream-related methods

        /// <summary>
        /// Creates a new instance of the <see cref="WmlDocument"/> class from
        /// this stream. Changes to the WmlDocument will not have any effect
        /// on this stream.
        /// </summary>
        /// <returns></returns>
        public static WmlDocument GetWmlDocument(this WordprocessingMemoryStream stream)
        {
            return new WmlDocument(stream.Path, stream);
        }

        #endregion WordprocessingMemoryStream-related methods

        #region SpreadsheetMemoryStream-related methods

        /// <summary>
        /// Creates a new instance of the <see cref="SmlDocument"/> class from
        /// this stream. Changes to the SmlDocument will not have any effect
        /// on this stream.
        /// </summary>
        /// <returns></returns>
        public static SmlDocument GetSmlDocument(this SpreadsheetMemoryStream stream)
        {
            return new SmlDocument(stream.Path, stream);
        }

        #endregion SpreadsheetMemoryStream-related methods

        #region PresentationMemoryStream-related methods

        /// <summary>
        /// Creates a new instance of the <see cref="PmlDocument"/> class from
        /// this stream. Changes to the PmlDocument will not have any effect
        /// on this stream.
        /// </summary>
        /// <returns></returns>
        public static PmlDocument GetSmlDocument(this PresentationMemoryStream stream)
        {
            return new PmlDocument(stream.Path, stream);
        }

        #endregion PresentationMemoryStream-related methods
    }
}
