# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.3.6](N/A) - 2018-04-08
#### Changed
* Update to TypeScript 2.8
* Update tsconfig.json with `noImplicitReturns: true` and `forceConsistentCasingInFileNames: true`, fix resulting issues
* Update package.json `dom-builder` and `xlsx-spec-models` to correct source url and fix require() paths
* Added tarball and package.json npm script `build-package` reference for creating tarball


--------
### [0.3.5](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/22cc05f1a48b963753f80d92d91b48f3745c9cab) - 2018-02-28
#### Changed
* Update to TypeScript 2.7
* Update dependencies: mocha, @types/chai, @types/mocha, @types/node
* enable tsconfig.json `noImplicitAny`


--------
### [0.3.4](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/e2f45a4a0aa55606a48283d35370ccaaf479e3af) - 2017-10-03
#### Changed
* Update dependency `xlsx-spec-models@0.4.0`


--------
### [0.3.3](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/6abe89b18ec226a52ee02ec983b60db4c52e5cfd) - 2017-08-24
#### Changed
* Fix `XlsxReaderWriter.saveXlsxFile()` writing data to the wrong files


--------
### [0.3.2](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/1f4446da1dba26be30894b8d53360fc88c041caa) - 2017-08-24
#### Changed
* Aditional changes to `XmlFileInst.newInst()` and `DocLikeFile` and `XmlDocFile` to make them proper classes


--------
### [0.3.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/afecaa263adf49bd2260e1702da1677099093b95) - 2017-08-24
#### Changed
* tsconfig.json `"noImplicitThis": true` and related code type changes

#### Fixed
* Serious bug in `XmlFileInst.newInst()` not setting up the returned instances correctly.


--------
### [0.3.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/ef364947128f20a38ba68c8ee1f2c284c4d0f4df) - 2017-08-18
#### Added
* `Workbook` file types and reader/writer added to `XlsxReaderWriter`
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
* `CellRefUtil.mergeCellSpans()` typo bug


--------
### [0.2.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/db984736c9c9e6d314404c10834c620d64ca7c21) - 2017-01-28
#### Added
* OpenXmlIo ReadWrite and ReadWriteNamed interfaces
* Added `XlsxDomErrorsImpl.expectNode()` to match new `dom-builder@0.2.0` definition
* `XlsxReaderWriter` support for `WorksheetDrawing` part (i.e. 'xl/drawings/drawing#.xml')
* `XlsxReaderWriter.loadXlsxFile()` new `loadSettings` parameter to allow caller to skip parsing various parts of the spreadsheet

#### Changed
* Major refactoring to simplify and denest interfaces, split open-xml reading/writing into seperate interfaces
  * `OpenXmlIo.ParsedFile` refactored into `OpenXmlIo.ReaderContext` and `OpenXmlIo.WriterContext` both of which now `extend DomBuilderHelper`
  * `XmlFileInst` now contains `readMulti()` and `writeMulti()` instead of nested `readOpenXml` and `writeOpenXml` properties
  * `XmlFileInst` now `extends DomBuilderHelper` instead of containing a 'domHelper' property
  * Renamed `ReadOpenXmlElement` -> `ElementsReader` and is a function interface rather than containing `readMulti` method definition
  * Renamed `WriteOpenXmlElement` -> `ElementsWriter` and is a function interface rather than containing `writeMulti` method definition
  * Renamed `XlsxReaderWriter` `loadExcelFileInst()` -> `loadXlsxFile()` and `saveExcelFileInst()` -> `saveXlsxFile()`

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