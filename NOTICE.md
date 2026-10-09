# Third-party renderers

The document-viewer integration code is MIT licensed. Dependency licenses are not replaced by that license. Builds include renderer license files under `licenses/` and preserve legal comments.

- PDF.js: Apache-2.0.
- docx-preview: Apache-2.0.
- @file-viewer/doc: MIT.
- @aiden0z/pptx-renderer: Apache-2.0.
- SheetJS Community Edition: Apache-2.0.
- cfb: Apache-2.0; generic OLE compound-file reading, not a presentation renderer.
- fflate: MIT; streaming DEFLATE decoding for worksheet windows.
- Local control icons: a small Lucide 1.52.0 SVG subset, including Feather-derived icons (ISC/MIT). Source: https://github.com/lucide-icons/lucide. Full attribution ships in `src/lucide-license.txt` and `licenses/Lucide.txt`; no icon runtime, CDN or additional dependency is loaded.

The legacy PPT parser, Worker and renderer are original project code implemented from public Microsoft MS-PPT/MS-ODRAW record definitions. No proprietary presentation runtime, watermark assets, protected engine source or reverse-engineered code is included. PDF.js resources are served locally. No document persistence is enabled.

The worksheet style resolver, formula interpreter and DOC page-layout reader are original implementations based on Open XML/MS-DOC specifications and documented Calc semantics. LibreOffice, ExcelJS and Univer are architectural references, not new bundled dependencies; their source and licenses are not replaced or incorporated by this notice.
