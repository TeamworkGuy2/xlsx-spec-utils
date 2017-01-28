TypeScript XLSX Model Utils
==============

TypeScript utilities for reading/writing OpenXML spreadsheet (XLSX) models from the [xlsx-spec-models](https://github.com/TeamworkGuy2/xlsx-spec-models) library.

Includes:
- File readers/writers for SpreadsheetML file formats (see the `/files/` directory).

- Utilities for creating cell 'A1' style references from row and column numbers (see `/utils/CellRefUtils.ts`).

- Utilities for working with styles and fonts, finding existing formats in a parsed spreadsheet, and creating new ones (see `/utils/`: `SharedStringsUtil.ts`, `StylesheetUtil.ts`, and `WorksheetUtil.ts`).
