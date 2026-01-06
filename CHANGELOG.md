## 1.2.0 (2026-01-06)

Features:
	- **Added XLSX export with border support**: New `borders` option enables exporting tables to XLSX format with borders around all cells
	- Added JSZip dependency for proper OpenXML XLSX generation
	- Added `excelFormat` option to choose between "xls" (legacy HTML-based) and "xlsx" (OpenXML) formats
	- XLSX export automatically enabled when `borders: true` is set
	- Updated demo to showcase border functionality with local library dependencies

Changes:
	- Added JSZip as a dependency for XLSX generation
	- Implemented OpenXML specification for Excel file generation
	- Added helper functions for generating proper XLSX file structure
	- Enhanced demo page with improved styling and border demonstration

Technical Details:
	- XLSX files follow OpenXML specification with proper structure: [Content_Types].xml, _rels/.rels, xl/workbook.xml, xl/styles.xml, and xl/worksheets/sheet1.xml
	- Borders are defined in styles.xml and applied via style references in worksheet cells
	- Maintains backward compatibility with legacy XLS export (default behavior unchanged)

## 1.1.2 (2019-02-06)

Changes:
	- Updated dependencies
    - Switched to yarn
    - merged PRs: https://github.com/rainabba/jquery-table2excel/pull/103,
         https://github.com/rainabba/jquery-table2excel/pull/110 and manually
         implemented changes from https://github.com/rainabba/jquery-table2excel/pull/108
    - Added 'serve' task for manual testing of library (rebuilds dist also)

## 1.1.1 (2017-04-28)

Changes:
	- Update to README.md and packaging

## 1.1.0 (2017-04-28)

Features:
	- Include td and th
	- Update dev dependencies and CONTRIBUTING.md
	- Cleanup repo tags/releases
