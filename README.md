# document-viewer

Browser-only, on-demand document previews for static sites and shared document URLs. Cloud builds require only Node.js and `npm ci`; no LibreOffice, Office installation or conversion backend is used.

## Supported formats

| Format | Rendering |
| --- | --- |
| PDF | PDF.js, page navigation and selectable text |
| DOCX | docx-preview, embedded pictures and pagination |
| DOC | MS-DOC browser parser |
| PPTX | Browser-native slide renderer |
| PPT | Project-owned MS-PPT parser and basic SVG/text preview, without a watermark |
| XLSX / XLS | SheetJS, worksheet switching and formatted cell values |

Office preview is not a pixel-perfect replacement for Office. Macros, linked external resources, scripts and spreadsheet formula evaluation are not enabled. Spreadsheet previews show up to 1,000 rows and 100 columns per worksheet. Password-protected files and unsupported document features show an error and preserve the original download link.

Legacy PPT uses this project's own parser, disposable Worker and renderer. The only container dependency is the Apache-2.0 `cfb` OLE reader, not a slide engine. No proprietary PPT engine or watermark assets are included. This implementation follows public Microsoft MS-PPT/MS-ODRAW definitions without copying protected engine code. See `NOTICE.md` for dependency licenses.

### Legacy PPT boundaries

This implementation follows the current persist-directory/edit chain and active slide order, excluding obsolete edits, notes and unused masters. It reads Unicode/compressed text, outline-text references, slide dimensions, basic shape positions, rectangles, ellipses, triangles, lines, explicit solid colors, rotation and embedded PNG/JPEG references. Parsing runs in a disposable Worker with bounded record/depth/slide counts. Slides have page navigation and selectable text.

It is **not** a complete replacement for a mature Office renderer: animations, charts, master/style inheritance, grouped-coordinate transforms, complete text-run formatting, gradients, arbitrary paths, OLE, EMF/WMF and encrypted PPT are unsupported. Each legacy PPT preview displays this limitation and preserves the original download. Extracted text without a supported positioned shape remains readable below the slide.

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

Load `styles.css` alongside the library. The default preview button performs the fetch; mounting performs no document or vendor requests. `viewer.close()` aborts fetches and releases rendering resources; `viewer.destroy()` also removes the component.

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
- Metadata resolution, authentication and share permissions belong to the calling app. Do not put secret tokens in static pages. Optional `previewSrc` selects an independently prepared preview while the download link keeps the original `src`.
- `maxBytes` defaults to 50 MiB. Office ZIP previews also enforce entry-count and expanded-size limits. `canLoad` is checked immediately before preview; the caller must call `close()` when consent or access is revoked.

## CSP and caching

Allow local modules/workers, local fetches, inline renderer styles, data/blob images and fonts, and `script-src 'wasm-unsafe-eval'` for WASM decoding. **Do not** add `unsafe-eval`. Remote files need only their explicitly approved origin in `connect-src`. Documents render in a script-disabled sandbox whose own CSP blocks external resources.

Build adapters can import `documentViewerAssets` from `@jsw-teams/document-viewer/assets`. It returns local asset descriptors without writing build output. Fingerprint the JS/CSS dependency graph, including `new URL()` worker references. PDF resources are copied to content-hashed directories; cache them for one year. Revalidate HTML, mutable document URLs and metadata. Never publicly cache authenticated responses.

## Tests

```sh
npm test
```

Tests use local browser fixtures. They do not upload documents to any service.
