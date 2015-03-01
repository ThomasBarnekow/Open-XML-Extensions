/*
 * PackagingExtensions.cs - Extensions for System.IO.Packaging classes
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
using System.Linq;
using System.Text;

using System.IO.Packaging.FlatOpc;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace System.IO.Packaging.Extensions
{
    /// <summary>
    /// This enum defines the Physical Package types.
    /// </summary>
    public enum PhysicalPackageType
    {
        Zip = 1,
        FlatOpc = 2
    }

    /// <summary>
    /// This class implements a Package factory. The methods in "Extended Package methods" 
    /// region could be additions to the <see cref="Package"/> class.
    /// </summary>
    public static class PackageFactory
    {
        // Default values for the Package.Open method overloads
        private static readonly FileMode _defaultFileMode = FileMode.OpenOrCreate;
        private static readonly FileAccess _defaultFileAccess = FileAccess.ReadWrite;
        private static readonly FileShare _defaultFileShare = FileShare.None;

        private static readonly FileMode _defaultStreamMode = FileMode.Open;
        private static readonly FileAccess _defaultStreamAccess = FileAccess.Read;

        #region Existing Package methods 

        #region File-based Open methods

        public static Package Open(string path)
        {
            return Package.Open(path, _defaultFileMode, _defaultFileAccess, _defaultFileShare);
        }

        public static Package Open(string path, FileMode packageMode)
        {
            return Package.Open(path, packageMode, _defaultFileAccess, _defaultFileShare);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess)
        {
            return Package.Open(path, packageMode, packageAccess, _defaultFileShare);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess, FileShare packageShare)
        {
            return Package.Open(path, packageMode, packageAccess, packageShare);
        }

        #endregion File-based Open methods

        #region Stream-based Open methods

        public static Package Open(Stream stream)
        {
            return Package.Open(stream, _defaultStreamMode, _defaultStreamAccess);
        }

        public static Package Open(Stream stream, FileMode packageMode)
        {
            return Package.Open(stream, packageMode, _defaultStreamAccess);
        }

        public static Package Open(Stream stream, FileMode packageMode, FileAccess packageAccess)
        {
            return Package.Open(stream, packageMode, packageAccess);
        }

        #endregion Stream-based Open methods
        
        #endregion Existing Package methods

        #region Extended Package methods

        #region File-based Open methods

        public static Package Open(string path, PhysicalPackageType packageType)
        {
            return Open(path, _defaultFileMode, _defaultFileAccess, _defaultFileShare, packageType);
        }

        public static Package Open(string path, FileMode packageMode, PhysicalPackageType packageType)
        {
            return Open(path, packageMode, _defaultFileAccess, _defaultFileShare, packageType);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess, PhysicalPackageType packageType)
        {
            return Open(path, packageMode, packageAccess, _defaultFileShare, packageType);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess, FileShare packageShare, PhysicalPackageType packageType)
        {
            if (packageType == PhysicalPackageType.Zip)
                return Package.Open(path, packageMode, packageAccess, packageShare);
            else
                return FlatOpcPackage.Open(path, packageMode, packageAccess, packageShare);
        }

        #endregion File-based Open methods

        #region Stream-based Open methods

        public static Package Open(Stream stream, PhysicalPackageType packageType)
        {
            return Open(stream, _defaultStreamMode, _defaultStreamAccess, packageType);
        }

        public static Package Open(Stream stream, FileMode packageMode, PhysicalPackageType packageType)
        {
            return Open(stream, packageMode, _defaultStreamAccess, packageType);
        }

        public static Package Open(Stream stream, FileMode packageMode, FileAccess packageAccess, PhysicalPackageType packageType)
        {
            if (packageType == PhysicalPackageType.Zip)
                return Package.Open(stream, packageMode, packageAccess);
            else
                return FlatOpcPackage.Open(stream, packageMode, packageAccess);
        }

        #endregion Stream-based Open methods

        #endregion Extended Package methods
    }

    /// <summary>
    /// This class demonstrates which static methods could be added to the 
    /// <see cref="WordprocessingDocument"/> class and its siblings.
    /// </summary>
    public static class WordprocessingDocumentFactory
    {
        #region Create

        #region Create on file

        public static WordprocessingDocument Create(string path, WordprocessingDocumentType type, PhysicalPackageType packageType)
        {
            return Create(path, type, true, packageType);
        }
        
        public static WordprocessingDocument Create(string path, WordprocessingDocumentType type, bool autoSave, PhysicalPackageType packageType)
        {
            Package package = null;
            if (packageType == PhysicalPackageType.FlatOpc)
                package = FlatOpcPackage.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            else
                package = Package.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);

            return WordprocessingDocument.Create(package, type, autoSave);
        }

        #endregion Create on file

        #region Create on Stream

        public static WordprocessingDocument Create(Stream stream, WordprocessingDocumentType type, PhysicalPackageType packageType)
        {
            return Create(stream, type, true, packageType);
        }

        public static WordprocessingDocument Create(Stream stream, WordprocessingDocumentType type, bool autoSave, PhysicalPackageType packageType)
        {
            Package package = null;
            if (packageType == PhysicalPackageType.FlatOpc)
                package = FlatOpcPackage.Open(stream, FileMode.Create, FileAccess.ReadWrite);
            else
                package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite);

            return WordprocessingDocument.Create(package, type, autoSave);
        }

        #endregion Create on Stream

        // We obviously don't need the following method:
        // public static WordprocessingDocument Create(Package package, WordprocessingDocumentType type, bool autoSave, PhysicalPackageType packageType)

        #endregion Create

        #region Open

        #region Open on file

        public static WordprocessingDocument Open(string path, bool isEditable, PhysicalPackageType packageType)
        {
            return Open(path, isEditable, new OpenSettings(), packageType);
        }
        
        public static WordprocessingDocument Open(string path, bool isEditable, OpenSettings openSettings, PhysicalPackageType packageType)
        {
            FileMode packageMode = isEditable ? FileMode.OpenOrCreate : FileMode.Open;
            FileAccess packageAccess = isEditable ? FileAccess.ReadWrite : FileAccess.Read;
            FileShare packageShare = isEditable ? FileShare.None : FileShare.Read;

            Package package = null;
            if (packageType == PhysicalPackageType.FlatOpc)
                package = FlatOpcPackage.Open(path, packageMode, packageAccess, packageShare);
            else
                package = Package.Open(path, packageMode, packageAccess, packageShare);

            return WordprocessingDocument.Open(package, openSettings);
        }

        #endregion Open on file

        #region Open on Stream

        public static WordprocessingDocument Open(Stream stream, bool isEditable, PhysicalPackageType packageType)
        {
            return Open(stream, isEditable, new OpenSettings(), packageType);
        }

        public static WordprocessingDocument Open(Stream stream, bool isEditable, OpenSettings openSettings, PhysicalPackageType packageType)
        {
            FileMode packageMode = isEditable ? FileMode.OpenOrCreate : FileMode.Open;
            FileAccess packageAccess = isEditable ? FileAccess.ReadWrite : FileAccess.Read;
            FileShare packageShare = isEditable ? FileShare.None : FileShare.Read;

            Package package = null;
            if (packageType == PhysicalPackageType.FlatOpc)
                package = FlatOpcPackage.Open(stream, packageMode, packageAccess);
            else
                package = Package.Open(stream, packageMode, packageAccess);

            return WordprocessingDocument.Open(package, openSettings);
        }

        #endregion Open on Stream

        // We obviously don't need the following method:
        // public static WordprocessingDocument Open(Package package, WordprocessingDocumentType type, bool autoSave, PhysicalPackageType packageType)

        #endregion Open
    }
}
