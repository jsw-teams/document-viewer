# document-viewer

Browser-only inline document previews for static sites and shared document URLs. Cloud builds require only Node.js and `npm ci`; no LibreOffice, Office installation or conversion backend is used.

## Supported formats

| Format | Rendering |
| --- | --- |
| PDF | PDF.js, continuous separated pages and selectable text |
| DOCX | docx-preview, embedded pictures and pagination |
| DOC | MS-DOC browser parser |
| PPTX | Browser-native slide renderer |
| PPT | Project-owned MS-PPT parser and basic SVG/text preview, without a watermark |
| XLSX / XLS | SheetJS, worksheet switching and formatted cell values |

Office preview is not a pixel-perfect replacement for Office. Macros, linked external resources, scripts and spreadsheet formula evaluation are not enabled. Spreadsheet previews show up to 1,000 rows and 100 columns per worksheet. Password-protected files and unsupported document features show an error. The viewer exposes no original-file download links.

Legacy PPT uses this project's own parser, disposable Worker and renderer. The only container dependency is the Apache-2.0 `cfb` OLE reader, not a slide engine. No proprietary PPT engine or watermark assets are included. This implementation follows public Microsoft MS-PPT/MS-ODRAW definitions without copying protected engine code. See `NOTICE.md` for dependency licenses.

### Legacy PPT boundaries

This implementation follows the current persist-directory/edit chain and active slide order, excluding obsolete edits, notes and unused masters. It reads Unicode/compressed text, outline-text references, slide dimensions, basic shape positions, rectangles, ellipses, triangles, lines, explicit solid colors, rotation and embedded PNG/JPEG references. Parsing runs in a disposable Worker with bounded record/depth/slide counts. Slides have continuous separated pages and selectable text. Both 16-bit and 32-bit MS-PPT client rectangles use top/left/right/bottom; MS-ODRAW child anchors use left/top/right/bottom. Local textbox text is retained alongside outline text. Text that cannot fit a positioned box remains readable below the slide instead of being clipped away.

Specifications: [MS-PPT RectStruct](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/8a58e3ae-2682-42d0-82cd-a41c2999584e), [SmallRectStruct](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/e47cb973-8480-4995-90b2-008bcb2ffc65) and [OfficeArtClientTextbox](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/f50070dd-a4dc-4edd-a446-c4fcc5c80ace).

The explicit `fUsefFitShapeToText` / `fFitShapeToText` flags in [MS-ODRAW Text Boolean Properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/ab9e4283-a47f-429c-8c2b-683a3a7f16d1) expand text boxes to measured text height within the slide. A thin stored anchor is not treated as a fixed clipping rectangle when these flags request autofit.

## Continuous pages

Documents scroll through separated, numbered pages rather than previous/next controls. Word uses authored page breaks, presentations use slide boundaries, and spreadsheets split the selected worksheet into 50-row pages while retaining worksheet tabs. PDF pages render near the scroll viewport and release distant rasters when the cache exceeds eight pages or 32 million pixels. The sandbox stays script-disabled; renderers and observers run in the parent page. Legacy Word HTML is parsed into an inert template and stripped of executable elements and event attributes before insertion.

It is **not** a complete replacement for a mature Office renderer: animations, charts, master/style inheritance, grouped-coordinate transforms, complete text-run formatting, gradients, arbitrary paths, OLE, EMF/WMF and encrypted PPT are unsupported. These technical boundaries are documented here rather than displayed as a warning in the viewer. Extracted text without a supported positioned shape remains readable below the slide.

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

Controls are locally styled, not native browser selectors. The component automatically follows the host's font, text, nearest opaque background and link colors, including sites without EdgePress theme variables. Common background/foreground/primary/surface variables are also recognized; `--document-viewer-*` variables allow explicit overrides. Accent text and focus colors are chosen for contrast rather than assuming every accent is dark. The sandbox synchronizes the palette when ancestor theme classes/styles, stylesheet content or the system color scheme change. Worksheet tables use the site palette, while Word, PDF and slide pages preserve their document colors.

Worksheet tabs support Left/Right, Home/End and roving keyboard focus, with selected state and a labeled panel. Tables expose row and column coordinates to assistive technology. Controls have 44 px minimum targets, visible focus and forced-colors support. Loading, errors and page positions are announced. Escape closes the preview even from inside the sandbox and restores focus to Preview. Integration tests check narrow layouts, light/dark palettes and WCAG contrast with axe; the host remains responsible for accessible theme colors.

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
- `maxBytes` defaults to 50 MiB. Office ZIP previews also enforce entry-count and expanded-size limits. `canLoad` is checked immediately before preview; the caller must call `close()` when consent or access is revoked.

## CSP and caching

Allow local modules/workers, local fetches, inline renderer styles, data/blob images and fonts, and `script-src 'wasm-unsafe-eval'` for WASM decoding. **Do not** add `unsafe-eval`. Remote files need only their explicitly approved origin in `connect-src`. Documents render in a script-disabled sandbox whose own CSP blocks external resources.

Build adapters can import `documentViewerAssets` from `@jsw-teams/document-viewer/assets`. It returns local asset descriptors without writing build output. Fingerprint the JS/CSS dependency graph, including `new URL()` worker references. PDF resources are copied to content-hashed directories; cache them for one year. Revalidate HTML, mutable document URLs and metadata. Never publicly cache authenticated responses.

## Tests

```sh
npm test
```

Tests use local browser fixtures. They do not upload documents to any service.
