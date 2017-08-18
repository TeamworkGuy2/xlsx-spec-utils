# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.3.0](N/A) - 2017-08-18
#### Added
* `Workbook` file types and reader/writer added to XlsxReaderWriter
* Added `XlsxFileType.getXmlFilePath()`
* Added `utils/RelationshipsUtil`

#### Changed
* Updated to TypeScript 2.4
* Updated to `xlsx-spec-models@0.3.0` which includes a `dom-builder@0.4.1` update
  * XmlFileInst is no longer a class, it is now a module with two sub-classes: `DocLikeFile` and `XmlDocFile`, use XmlFileInst.newInst() exclusively instead of calling `new XmlFileInst()` as a constructor
  * XmlFileInst.newInst() now accepts `DocumentLike` as well as `XMLDocument`
  * XmlFileInst.writeMulti() now returns `ElementLike` instead of `HTMLElement`

#### Removed
* Moved `open-xml-io.d.ts` from this project to the `xlsx-spec-models` library


--------
### [0.2.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/7030331800214fcbc056606c43df722585d07276) - 2017-05-09
#### Added
* Updated to TypeScript 2.3, add tsconfig.json, use @types/ definitions

#### Fixed
* CellRefUtil.mergeCellSpans() typo bug


--------
### [0.2.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/db984736c9c9e6d314404c10834c620d64ca7c21) - 2017-01-28
#### Added
* OpenXmlIo ReadWrite and ReadWriteNamed interfaces
* Added XlsxDomErrorsImpl.expectNode() to match new `dom-builder@0.2.0` definition
* XlsxReaderWriter support for WorksheetDrawing part (i.e. 'xl/drawings/drawing#.xml')
* XlsxReaderWriter.loadXlsxFile() new `loadSettings` parameter to allow caller to skip parsing various parts of the spreadsheet

#### Changed
* Major refactoring to simplify and denest interfaces, split open-xml reading/writing into seperate interfaces
  * OpenXmlIo.ParsedFile refactored into OpenXmlIo.ReaderContext and OpenXmlIo.WriterContext both of which now extend DomBuilderHelper
  * XmlFileInst now contains readMulti() and writeMulti() instead of nested 'readOpenXml' and 'writeOpenXml' properties
  * XmlFileInst now directly extends DomBuilderHelper instead of having 'domHelper' property
  * Renamed ReadOpenXmlElement -> ElementsReader and is a function interface rather than containing 'readMulti' method definition
  * WriteOpenXmlElement -> ElementsWriter and is a function interface rather than containing 'writeMulti' method definition
  * Renamed XlsxReaderWriter loadExcelFileInst() -> loadXlsxFile() and saveExcelFileInst() -> saveXlsxFile()

#### Removed
* Removed Read, ReadNamed, Write, WriteNamed interfaces, see new ReadWrite and ReadWriteNamed interfaces


--------
### [0.1.3](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/a6549b833a80912d52724a9ed3074a8865e8e884) - 2016-11-12
#### Added
Added missing documentation to the `/files/` classes, `open-xml-io.d.ts`, and `StylesheetUtil.ts`

#### Fixed
`WorksheetUtil.createCellSimpleFormula()` incorrectly picking `cell.val` when `cell.formulaString` was non-null


--------
### [0.1.2](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/3153a109a74c2ddaeada238c176f43ba648657a4) - 2016-08-24
#### Fixed
XlsxReaderWriter skips reading and writing optional spreadsheet parts, such as calcChain, sharedStrings, and comments instead of throwing an error


--------
### [0.1.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/9aee05563241ee8898d6373e9f95017d2f78f8fe) - 2016-05-30
#### Changed
Fixed to use renamed dom-builder library (was previously ts-dom-builder)


--------
### [0.1.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/b521f1c9ef97afcbd63d1cbaf4cd3ec028670beb) - 2016-05-28
#### Added
Initial commit of XLSX read/write utils.