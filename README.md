# Open-XML-Extensions

This project provides a number of extensions to:

* the **Open XML SDK 2.5** developed by Microsoft and now maintained by Eric White (see https://github.com/OfficeDev/Open-XML-SDK) and
* the **PowerTools for Open XML** developed by Eric White (see http://powertools.codeplex.com).

Currently, this project consists of the following solutions:

* **OpenXmlExtensions**: MemoryStreams, Transforms, and Extensions for Open XML Documents; and
* **OpenXmlPowerTools**: Proposed enhancements of Eric White's PowerTools for Open XML.

## OpenXmlExtensions

### DocumentFormat.OpenXml.IO: OpenXmlMemoryStreams

MemoryStreams are the preferred way to keep Open XML Packages, e.g., WordprocessingDocument instances, in memory for various changes or transformations performed on them. As part of the *DocumentFormat.OpenXml.IO* project, this solution provides a number of classes derived from MemoryStream that contain a number of useful additional features for working with Open XML Documents. All of them provide a certain level of "type safety", much like their corresponding Open XML SDK classes.

* **OpenXmlMemoryStream**: Abstract base class derived from MemoryStream. Corresponds to OpenXmlPackage. Includes Save and SaveAs methods, among other things, for saving an OpenXmlMemoryStream to a file.
* **WordprocessingMemoryStream**: Derived from OpenXmlMemoryStream, this class is the companion for WordprocessingDocuments. It includes constructors for creating WordprocessingMemoryStreams from byte arrays, files, and streams. It further includes a static method for creating a "minimum" WordprocessingDocument as a starting point for document generation. Lastly, it includes methods for opening a WordprocessingDocument from the stream. 
* **SpreadsheetMemoryStream**: Derived from OpenXmlMemoryStream, this is the companion for SpreadsheetDocuments and provides the same general features as WordprocessingMemoryStream. 
* **PresentationMemoryStream**: Derived from OpenXmlMemoryStream, this is the companion for PresentationDocuments and provides the same general features as WordprocessingMemoryStream and SpreadsheetMemoryStream.

### DocumentFormat.OpenXml.Transforms: OpenXmlTransforms

Documentation Work in Progress ...

### Further Extensions

Documentation Work in Progress ...

## OpenXmlPowerTools

This solution contains Eric White's PowerTools for Open XML (see http://powertools.codeplex.com) release 2.7.04 with some proposed fixes. The objective is to have these included in future releases of the PowerTools for Open XML.

The solution contains two static utility classes related to the OpenXmlMemoryStream classes described above:

* **PtMemoryStreamExtensions**: Provides extension methods for creating OpenXmlPowerToolsDocument, WmlDocument, SmlDocument, and PmlDocument instances from the corresponding OpenXmlMemoryStreams.
* **PtMemoryStreamFactory**: Provides static factory methods for creating a WordprocessingMemoryStream, SpreadsheetMemoryStream, and PresentationMemoryStream from a WmlDocument, SmlDocument, and PmlDocument, respectively. 