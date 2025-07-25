# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [1.0.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/releases/tag/v1.0.0) - 2025-07-25
#### Changed
Time to mark this package stable and v1!

* BREAKING: update exports to use standard named exports rather than a default export object. This requires updating import statements from `import XlsxReaderWriter from 'xlsx-spec-utils/XlsxReaderWriter';` to `import { XlsxReaderWriter } from 'xlsx-spec-utils/XlsxReaderWriter';`, notice the parenthesis now required around the import type.
* Update dependency `xlsx-spec-models` to `v1.0.0`


--------
### [0.8.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/releases/tag/v0.8.1) - 2025-07-25
#### Fixed
* `package.json` version


--------
### [0.8.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/releases/tag/v0.8.0) - 2024-01-05
#### Changed
* Update to TypeScript 4.9


--------
### [0.7.2](https://github.com/TeamworkGuy2/xlsx-spec-utils/releases/tag/v0.7.2) - 2023-12-07
#### Changed
* Build: Enable TypeScript `strict` compile option
* Build: rename `tsc` npm command in package.json to `build`


--------
### [0.7.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/8bd31e54643f1234b1b54ef1f28ae114bfe9d08c) - 2022-02-21
#### Changed
* package.json dependency update to use npm version of `xlsx-spec-models@0.8.2` instead of github release tarball


--------
### [0.7.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/f27f2d1eb02e4e42786056476c28063642e29f4a) - 2022-01-03
#### Changed
* Update to TypeScript 4.4


--------
### [0.6.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/6853144b23085ac55caaa45cbdf3034f63784f29) - 2021-06-12
#### Changed
* Update to TypeScript 4.3


--------
### [0.5.1](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/1b330bf037aa68e6cca45b3aa8f5b4ff69a5dfb4) - 2021-02-12
#### Fixed
* `XmlFileReadWriter.saveToDom()` regression error introduced in `0.5.0`
* Fix an error message typo


--------
### [0.5.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/820c4c6c33338ea023b6428d9abf352273391823) - 2021-01-01
#### Change
* TypeScript - enable `strict` compilation
  * Fix compile errors related to `strict`
* Update dependency `dom-builder@0.9.0` (API refactor and better attribute reading/writing) and `xlsx-spec-models@0.6.0`


--------
### [0.4.0](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/ddd503285f10b08f68620b254cd0aeaa98dbabbd) - 2020-09-05
#### Change
* Update to TypeScript 4.0


--------
### [0.3.14](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/efe4d34a91150842bdd6a6d70e2b195e27549150) - 2019-11-08
#### Change
* Update to TypeScript 3.7


--------
### [0.3.13](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/7ef267397d4acb557afd6e28114168d3de61a194) - 2019-07-06
#### Change
* Update to TypeScript 3.5


--------
### [0.3.12](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/4fe468d5539d8a77671d0dd3a47e41db607287f1) - 2019-05-24
#### Change
* `dom-builder` dependency update to v0.7.0 (improved attribute handling)
* `xlsx-spec-models@0.4.8` dependency update (to use `dom-builder@0.7.0`)


--------
### [0.3.11](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/daf62f5886dacfb0101a989e9bb2e49c80fc5d0f) - 2019-03-21
#### Fixed
* `dom-builder` import/reference paths not being updated to `@twg2/dom-builder`


--------
### [0.3.10](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/ab5f36d023f8871dafc2b611099fef04c11aa3ad) - 2019-03-21
#### Changed
* Switch `dom-builder` dependency from github to npm `@twg2/dom-builder`, update 'xlsx-spec-models@0.4.6` to also use `@twg2/dom-builder`


--------
### [0.3.9](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/0272d825acff1f21068f0555c79c6c1dbae22fa3) - 2018-12-29
#### Changed
* Update to TypeScript 3.2
* Update @types/ dependencies


--------
### [0.3.8](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/e6f21a7627b63b6ab20455a95dee5caf07ad1ea8) - 2018-10-20
#### Changed
* Switch `package.json` github dependencies from tag urls to release tarballs to simplify npm install (doesn't require git to npm install tarballs)
* Added `repository` to `package.json`


--------
### [0.3.7](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/4400862ebc360662fe4ac6fe1b71b9a98d5647c0) - 2018-10-17
#### Changed
* Update to TypeScript 3.1
* Update dev dependencies and @types
* Enable `tsconfig.json` `strict` and fix compile errors
* Removed compiled bin tarball in favor of git tags


--------
### [0.3.6](https://github.com/TeamworkGuy2/xlsx-spec-utils/commit/8ee37b2b6d298bd06c587dac94ad3137f6b4231e) - 2018-04-08
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
* Additional changes to `XmlFileInst.newInst()` and `DocLikeFile` and `XmlDocFile` to make them proper classes


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