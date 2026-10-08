# document-viewer

Curious about improving this project? Vibe Coding and AI-assisted contributions are welcome, with no tool restrictions. Start with [Contributing](CONTRIBUTING.md), follow [AGENTS.md](AGENTS.md), and share a small, understandable change with reproducible tests. [Report a bug or idea](https://github.com/jsw-teams/document-viewer/issues/new/choose) · [Security](SECURITY.md) · [License](LICENSE).

Browser-only inline document previews for static sites and shared document URLs. Cloud builds require only Node.js and `npm ci`; no LibreOffice, Office installation or conversion backend is used.

## Supported formats

| Format | Rendering |
| --- | --- |
| PDF | PDF.js, per-page CropBox/rotation, selectable text and zoom-aware raster rendering |
| DOCX | docx-preview, authored page geometry, saved page breaks, headers, footers and notes |
| DOC | MS-DOC browser parser plus source FIB/SEPX paper geometry |
| PPTX | Browser-native slide renderer |
| PPT | Project-owned MS-PPT parser, positioned shapes, character/paragraph runs and master text styles |
| XLSX / XLS | Windowed worksheets, accessible tabs, cell addresses, formats and bounded formula previews |

Preview fidelity is verified with format-specific layout fixtures and real documents, rather than inferred from text extraction alone. Macros, linked external resources and document scripts never execute. Saved spreadsheet results are preferred; missing results use a bounded, original interpreter with Calc-compatible operator semantics. Worksheets are not truncated to an arbitrary number of rows or columns: scroll or enter a cell address to visit the complete used range. The viewer exposes no editing or original-file download controls.

Legacy PPT uses this project's own parser, disposable Worker and renderer. The only container dependency is the Apache-2.0 `cfb` OLE reader, not a slide engine. No proprietary PPT engine or watermark assets are included. This implementation follows public Microsoft MS-PPT/MS-ODRAW definitions without copying protected engine code. See `NOTICE.md` for dependency licenses.

### Presentation reconstruction

This implementation follows the current persist-directory/edit chain and active slide order, excluding obsolete edits, notes and unused masters. It reads Unicode/compressed text, outline-text references, slide dimensions, basic shape positions, rectangles, ellipses, triangles, lines, explicit solid colors, rotation and embedded PNG/JPEG references. Parsing runs in a disposable Worker with bounded record/depth/slide counts. Slides have continuous separated pages and selectable text. Both 16-bit and 32-bit MS-PPT client rectangles use top/left/right/bottom; MS-ODRAW child anchors use left/top/right/bottom. Local textbox text is retained alongside outline text. Text that cannot fit a positioned box remains readable below the slide instead of being clipped away.

Text restoration reads MS-PPT paragraph and character runs: font references, point sizes, bold/italic/underline, literal and scheme colors, alignment, spacing, indentation, bullets and direction. Placeholder runs inherit main-master text levels, with document defaults and explicit overrides. Malformed style records do not erase readable text. PPTX retains the existing renderer's slide model and uses its adaptive resizing instead of freezing the initial viewport width.

Specifications: [MS-PPT RectStruct](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/8a58e3ae-2682-42d0-82cd-a41c2999584e), [SmallRectStruct](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/e47cb973-8480-4995-90b2-008bcb2ffc65), [TextCFException](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/c75024a2-14cb-4d7d-9964-bdab2fcd9d93) and [TextMasterStyleAtom](https://learn.microsoft.com/en-au/openspecs/office_file_formats/ms-ppt/5febad27-0c48-4f98-b655-562b986f5874).

The explicit `fUsefFitShapeToText` / `fFitShapeToText` flags in [MS-ODRAW Text Boolean Properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/ab9e4283-a47f-429c-8c2b-683a3a7f16d1) expand text boxes to measured text height within the slide. A thin stored anchor is not treated as a fixed clipping rectangle when these flags request autofit.

## Continuous pages

Documents scroll through separated, numbered pages rather than previous/next controls. Word uses authored page breaks, presentations use slide boundaries, and spreadsheets split the selected worksheet into 50-row pages. The read-only Office-inspired chrome provides the document title once, a format marker, zoom/fit controls and page position. A styled expanded workspace retains the original iframe and resources, traps keyboard focus and returns to the inline view with Escape. Worksheets have a bottom tab strip, a cell-address box and a value/formula bar; selecting a formula displays its expression without enabling editing.

DOCX retains the source paper width, margins and table layout, including centered tables, saved page breaks, headers, footers, footnotes and endnotes. Legacy DOC reads the first section's authored width, height and margins from the FIB section directory and SEPX properties. Word initially scales paper to fit narrow viewports without changing its layout; Fit width follows workspace resizing. Worksheet styles are isolated from Word tables. Worksheet pages grow to their rendered contents, so short two-column sheets are not clipped by a fixed page height. Columns use stored widths; Open XML character widths use a seven-pixel maximum-digit fallback, so unavailable or differently sized source fonts can still affect exact fidelity.

PDF requests 64 KiB HTTP byte ranges when the source supports them and exposes Content-Range through CORS. Automatic full-file prefetch is disabled. Each placeholder uses its own CropBox and rotation. Zoom and workspace resizing cancel stale paint work, discard old rasters and repaint nearby pages at the requested pixel density, including saved annotation appearances. Distant rasters and text layers are removed immediately, canvas backing stores are cleared on success or cancellation, and the active raster budget stays at eight pages / 32 million pixels. Unsupported range servers fall back to a bounded full read. PPTX uses the upstream renderer's windowed mount/unmount lifecycle with limited overscan. Word and legacy PPT discard distant DOM and decoded images; compact gzip page representations restore content when revisited, without retaining hidden DOM trees. The sandbox stays script-disabled, including restored pages.

XLSX is scanned in a disposable Worker using streaming DEFLATE/XML, retaining only requested cells and the shared strings those cells reference. No complete expanded worksheet or shared-string table is cached. The UI retains at most five 50-row/20-column windows; horizontal scrolling and cell addresses can reach the remaining columns. It restores theme/indexed colors and HSL tint, theme fonts, rich runs, borders, fills, alignment, row heights, hidden rows/columns, inherited row/column styles, numeric/date formats and merged cells. Merged cells crossing a window retain their source anchor. Switching worksheets disposes the previous window. XLS uses SheetJS to parse only the selected sheet inside the Worker and discards that parsed workbook after returning the requested window.

Legacy DOC parsing and section-property decoding also run in a disposable Worker. Source buffers are transferred rather than copied into the parent context; closing terminates in-progress parsing, and the Worker is discarded once sanitized HTML/CSS can be laid out in the parent-controlled sandbox.

### Formula previews

The project-owned interpreter follows [LibreOffice Calc operator precedence](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060199.html), including left-associative exponentiation. It supports arithmetic, comparisons, concatenation, relative/absolute references, cross-sheet ranges, named ranges and shared-formula rebasing. Implemented functions include SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, PRODUCT, SUMIF, COUNTIF, AVERAGEIF, IF, IFERROR, AND, OR, NOT, TRUE, FALSE, ABS, INT, SQRT, ROUND, MOD, POWER, LEN, LOWER, UPPER and TRIM. Conditional branches are lazy; saved results remain authoritative. Missing results show real calculation errors rather than disappearing.

Evaluation never compiles JavaScript, executes macros or invokes networking functions. Dependency blocks are capped at four 256-row/32-column blocks, computed results at 8,192 entries, and all calculation caches are discarded after each window. Work is cancellable and bounded by expression depth and operation count. Unknown functions return `#NAME?`; circular references return `#REF!`. This is an original preview interpreter, not a bundled LibreOffice engine or a claim of its complete function set.

These are rendering-memory optimizations, not a claim of constant total memory for every format. Compressed Office source bytes, metadata and compact page representations remain necessary to revisit content. Word/PPTX still require an initial parse/layout; legacy XLS can have a significant per-sheet parsing peak. Closing aborts requests, terminates Workers and releases all component-owned resources. Malformed archives, huge individual XML elements and decompression bombs still have safety budgets; those budgets must not be confused with truncating legitimate worksheet ranges.

Further reconstruction work needs focused fixtures for nested title-master chains, grouped legacy-PPT transforms, gradients, arbitrary paths, charts/OLE/metafiles, legacy-DOC multi-section mapping and overflow pagination. These are development targets, not reasons to skip ordinary document layout or hide text. Password-protected files require a separate access design. No claim of complete pixel parity with Office or LibreOffice is made by the current tests.

Design references include [LibreOffice Calc](https://docs.libreoffice.org/sc.html), [ExcelJS](https://github.com/exceljs/exceljs), [Univer](https://github.com/dream-num/univer) and the existing DOCX/PPTX/PDF dependencies. Standards-based parsing and rendering here are original code; no additional conversion service or third-party UI framework is introduced.

## Build and integrate

```sh
npm ci
npm run build
```

Deploy the complete `dist/` directory, not just its entry module. Heavy renderers and fonts load only after preview is requested for their format. Serve `.mjs` / `.js` as JavaScript and `.wasm` as `application/wasm`; never return an HTML fallback for these resources.

```js
import { mountDocument } from './document-viewer/index.js';

const viewer = mountDocument(document.querySelector('#preview'), {
  src: '/documents/report.docx',
  title: 'Annual report',
  locale: 'zh-CN'
});
```

Load `styles.css` alongside the library. Mounting opens the document automatically without moving keyboard focus, subject to `canLoad`. Set `autoOpen: false` to require explicit activation. After closing, the Preview button reopens the document. `viewer.close()` aborts fetches and releases rendering resources; `viewer.destroy()` also removes the component. Removing download links is a UI choice, not access control: a browser must receive the document bytes to render them.

### Theme and accessibility

Controls are locally styled, not native browser selectors. The component automatically follows the host's font, text, nearest opaque background and link colors, including sites without EdgePress theme variables. Common background/foreground/primary/surface variables are also recognized; `--document-viewer-*` variables allow explicit overrides. Accent text and focus colors are chosen for contrast rather than assuming every accent is dark. The sandbox synchronizes the palette when ancestor theme classes/styles, stylesheet content or the system color scheme change. Worksheet chrome and row/column headers use the site palette, while document cells, Word, PDF and slide pages preserve their document colors.

Worksheet tabs support Left/Right, Home/End and roving keyboard focus, with selected state and a labeled panel. Tables expose row and column coordinates to assistive technology. Controls have 44 px minimum targets, visible focus and forced-colors support. Loading, errors and page positions are announced. Escape closes the preview even from inside the sandbox and restores focus to Preview. Integration tests check narrow layouts, light/dark palettes and WCAG contrast with axe; the host remains responsible for accessible theme colors.

DOCX Symbol list bullets encoded as U+F0B7 are rendered with the Unicode U+2022 equivalent, so previews do not depend on an installed proprietary Symbol font. Unrelated fonts and private-use glyphs remain unchanged. The equivalence follows the [Unicode Consortium's Symbol encoding mapping](https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt).

## Shared URLs

```js
mountDocument(container, {
  src: sharedDocument.downloadUrl,
  format: 'xlsx',
  title: sharedDocument.name,
  canLoad: () => callerHasCurrentConsent()
});
```

- `src` must be a direct file URL, not a sharing HTML page. Extensionless URLs require `format`: `pdf`, `doc`, `docx`, `ppt`, `pptx`, `xls` or `xlsx`.
- The file server must allow CORS for the viewing site's origin. No proxy is created to bypass CORS. Fetches omit credentials and referrers; redirects are rejected to prevent an unapproved origin receiving a signed URL. Resolve sharing links to their authorized final download URL through the caller's consent-controlled integration.
- Metadata resolution, authentication and share permissions belong to the calling app. Do not put secret tokens in static pages. Optional `previewSrc` selects an independently prepared preview; neither URL is exposed as a download link.
- `maxBytes` defaults to 50 MiB for Office and 1 GiB for range-served PDF. PDF servers without usable ranges remain bounded to a 50 MiB full read. Office ZIP previews enforce entry-count and expanded-size safety budgets; streamed XLSX parts permit larger expanded sheets without allocating their complete XML. `canLoad` is checked immediately before preview; the caller must call `close()` when consent or access is revoked.

## CSP and caching

Allow local modules/workers, local fetches, inline renderer styles, data/blob images and fonts, and `script-src 'wasm-unsafe-eval'` for WASM decoding. **Do not** add `unsafe-eval`. Remote files need only their explicitly approved origin in `connect-src`. Documents render in a script-disabled sandbox whose own CSP blocks external resources.

Build adapters can import `documentViewerAssets` from `@jsw-teams/document-viewer/assets`. It returns local asset descriptors without writing build output. Fingerprint the JS/CSS dependency graph, including `new URL()` worker references. PDF resources are copied to content-hashed directories; cache them for one year. Revalidate HTML, mutable document URLs and metadata. Never publicly cache authenticated responses.

## Tests

```sh
npm test
```

Tests use local browser fixtures. They do not upload documents to any service.
