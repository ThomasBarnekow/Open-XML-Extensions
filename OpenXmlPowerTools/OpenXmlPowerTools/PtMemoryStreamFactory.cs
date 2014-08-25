/*
 * PtMemoryStreamFactiry.cs - PowerTools-related Factory for MemoryStreams
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

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Utility class providing a selection of MemoryStream-related tools.
    /// </summary>
    public static class PtMemoryStreamFactory
    {
        #region WordprocessingMemoryStream-related methods

        /// <summary>
        /// Creates a new <see cref="WordprocessingMemoryStream"/> instance, copying
        /// the <see cref="WmlDocument"/>'s byte array and using the file name.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static WordprocessingMemoryStream CreateWordprocessingMemoryStream(WmlDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return new WordprocessingMemoryStream(document.DocumentByteArray, document.FileName);
        }

        #endregion WordprocessingMemoryStream-related methods

        #region SpreadsheetMemoryStream-related methods

        /// <summary>
        /// Creates a new <see cref="SpreadsheetMemoryStream"/> instance, copying
        /// the <see cref="SmlDocument"/>'s byte array and using the file name.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static SpreadsheetMemoryStream CreateSpreadsheetMemoryStream(SmlDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return new SpreadsheetMemoryStream(document.DocumentByteArray, document.FileName);
        }

        #endregion SpreadsheetMemoryStream-related methods

        #region PresentationMemoryStream-related methods

        /// <summary>
        /// Creates a new <see cref="PresentationMemoryStream"/> instance, copying
        /// the <see cref="PmlDocument"/>'s byte array and using the file name.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static PresentationMemoryStream CreatePresentationMemoryStream(PmlDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            return new PresentationMemoryStream(document.DocumentByteArray, document.FileName);
        }

        #endregion SpreadsheetMemoryStream-related methods
    }
}
