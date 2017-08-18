TypeScript XLSX Model Utils
==============

TypeScript utilities for reading/writing OpenXML spreadsheet (XLSX) models from the [xlsx-spec-models](https://github.com/TeamworkGuy2/xlsx-spec-models) library.

Includes:
- File readers/writers for SpreadsheetML file formats (see the `/files/` directory).

- Utilities for creating cell 'A1' style references from row and column numbers (see `/utils/CellRefUtils.ts`).

- Utilities for working with styles and fonts, finding existing formats in a parsed spreadsheet, and creating new ones (see `/utils/`: `SharedStringsUtil.ts`, `StylesheetUtil.ts`, and `WorksheetUtil.ts`).

Example:
```ts
var jszip = // JSZip constructor function: 'new jszip(Uint8Array)' which returns a JSZip unzipped file data structure (which is not actually used by xlsx-spec-utils), see below

// Load an existing file as a template
var excelDataUnzipped = XlsxReaderWriter.readZip(/*Uint8Array*/excelZippedFileData, jszip);
var workbook = XlsxReaderWriter.loadXlsxFile({ }, (path) => excelDataUnzipped.files[path] != null ? excelDataUnzipped.files[path].asText() : null);

// Create a new, blank xlsx file
var workbook = XlsxReaderWriter.loadXlsxFile({ }, (path) => XlsxReaderWriter.defaultFileCreator(path));

// Write the xlsx file data back to the JSZip instance
XlsxReaderWriter.saveXlsxFile(workbook, (path, data) => excelDataUnzipped.file(path, data));
```