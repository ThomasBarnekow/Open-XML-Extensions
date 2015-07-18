# Open-XML-Extensions

This project provides a number of extensions to the **Open XML SDK 2.6** developed by
Microsoft and maintained by Eric White (see https://github.com/OfficeDev/Open-XML-SDK).
The Open XML SDK also includes extensions by myself, e.g., saving and cloning, Flat
OPC conversion, and document creation from templates.

## Current Project Status

<a href="https://scan.coverity.com/projects/5787">
  <img alt="Coverity Scan Build Status"
       src="https://scan.coverity.com/projects/5787/badge.svg"/>
</a>

## ContractArchitect.OpenXml.Transformation

Functional transformations are an interesting way to implement XML-based features.
The `OpenXmlTransformation<TDocument>` class and its subclasses contained in the 
`ContractArchitect.OpenXml.Transformation` namespace provide a framework for 
implementing such transformations:

* `OpenXmlTransformation<TDocument>`: Abstract base class currently defining methods for
  transforming Flat OPC `string`s, Flat OPC `XDocument`s, and `WordprocessingDocument`s
  (`SpreadsheetDocument`s and `PresentationDocument`s to follow).

* `FlatOpcStringTransformation<TDocument>`: Abstract base class that is subclassed by concrete
  transformations implementing a `string`-based transformation. This class automatically
  translate other supported inputs, i.e., `XDocument` and `WordprocessingDocument`
  (at this time) into a `string` and the `string` transformed by `Transform(string)`
  back into the original format.

* `FlatOpcDocumentTransformation<TDocument>`: Abstract base class that is subclassed by concrete
  transformations implementing an `XDocument`-based transformation.
  Like `FlatOpcStringTransformation<TDocument>`, this class translates between formats.

* `OpenXmlPackageTransformation<TDocument>`: Abstract base class that should be subclassed
  by concrete transformations that perform their specific operation on instances of
  `OpenXmlPackage` or, more specifically, instances of its subclasses. This class translates
  between formats.

* `WordprocessingDocumentTransformation`: Abstract base class that is subclassed by
  concrete transformations implementing a `WordprocessingDocument`-based transformation.
  Again, the class translates between formats.

* `XslOpenXmlTransformation<TDocument>`: This class is derived from 
  `FlatOpcStringTransformation<TDocument>` and uses an XSL stylesheet to perform the
  actual transformation.

Using the framework, you can basically write transforms the way you like. Are you a
Linq-to-XML person? Derive from `FlatOpcDocumentTransformation<TDocument>` and implement

* `Transform(XDocument)` to produce a Flat OPC `XDocument`

while callers might use

* `Transform(string)` to transform `string` representations or 
* `Transform(WordprocessingDocument)` to transform `WordprocessingDocument`s

without really caring about which way you implemented the actual transformation.
Do you want to leverage your XSL stylsheets? Go ahead and use `XslOpenXmlTransformation`.
You're the Open XML SDK programmer? Go ahead and derive from 
`WordprocessingDocumentTransformation` and use the strongly typed API. 

## ContractArchitect.OpenXml.Extensions

This namespace contains one class for each Open XML SDK class for which I've implemented
various extensions to make my life a little easier.
