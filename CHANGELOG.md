# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.1.3](N/A) - 2016-11-12
#### Added
Added missing documentation to the `/files/` classes, `open-xml-io.d.ts`, and `StylesheetUtil.ts'

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