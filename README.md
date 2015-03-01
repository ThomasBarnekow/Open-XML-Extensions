# Open-XML-Extensions

This project provides a number of extensions to the **Open XML SDK 2.5** developed by Microsoft and now maintained by Eric White (see https://github.com/OfficeDev/Open-XML-SDK).

## DocumentFormat.OpenXml.Transforms: OpenXmlTransforms

Functional transforms are an interesting way to implement XML-based features. The OpenXmlTransform class and its subclasses contained in the DocumentFormat.OpenXml.Transforms namespace provide a framework for implementing such transforms:

* **OpenXmlTransform**: Abstract base class currently defining methods for transforming Flat OPC strings, Flat OPC XDocuments, and WordprocessingDocuments (SpreadsheetDocuments and PresentationDocuments to follow). 
* **FlatOpcStringTransform**: Abstract base class that is subclassed by concrete transforms implementing a string-based transform. This class provides translators which translate other supported inputs, i.e., XDocument and WordprocessingDocument (at this time) into a string and the string transformed by string Transform(string) back into the original format. 
* **FlatOpcDocumentTransform**: Abstract base class that is subclassed by concrete transforms implementing an XDocument-based transform. Like FlatOpcString transform, this class also translates between formats.
* **WordprocessingDocumentTransform**: Abstract base class that is subclassed by concrete transforms implementing a WordprocessingDocument-based transform. Again, the class translates between formats.
* **XslOpenXmlTransform**: This class is derived from FlatOpcStringTransform and uses an XSL stylesheet to perform the actual transform.

Using the framework, you can basically write transforms the way you like. Are you a Linq-to-XML person? Derive from FlatOpcDocumentTransform and implement 

* XDocument Transform(XDocument)

while callers might use

* string Transform(string) or 
* WordprocessingDocument Transform(WordprocessingDocument)

without really caring about which API you used to implement the actual transform. Do you want to leverage your XSL stylsheets? Go ahead and use XslOpenXmlTransform. You're the Open XML SDK programmer? Go ahead and derive from WordprocessingDocumentTransform and use the strongly typed API. 

Things I'll add next will include support for the other kinds of OpenXml packages and complex transforms (e.g., chains or pipelines). 

## DocumentFormat.OpenXml.Extensions: Miscellaneous extensions

This namespace contains one class for each Open XML SDK class for which I've implemented various extensions to make my life a little easier.

## System.IO.Packaging.FlatOpc: FlatOpcPackages

This namespace contains an exploratory prototype of the FlatOpcPackage class, a subclass of System.IO.Packaging.Package that can be used in conjunction with WordprocessingDocument and its siblings to store documents in Flat OPC format right away. It uses the FlatOpcPackagePart and FlatOpcPackagePartStream classes to do its job. 
