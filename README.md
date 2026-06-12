# Word to Markdown

A browser-only `.docx` to Markdown converter designed for GitHub Pages.

## Features

- Converts Word `.docx` files to Markdown in the browser
- No backend and no upload step
- Supports headings, paragraphs, links, lists, tables, bold text and italics
- Provides editable Markdown output, a basic preview, clipboard copy and `.md` download
- Works as a static GitHub Pages site

## Publish with GitHub Pages

1. Create a public GitHub repository named `word-to-markdown`.
2. Upload the files in this folder to the repository root.
3. Open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select branch **main** and folder **/(root)**, then save.

The site will be available at:

```text
https://YOUR-GITHUB-USERNAME.github.io/word-to-markdown/
```

## Run locally

Open `index.html` in a browser.

## Privacy

Conversion happens locally in the browser. The document is not sent to a server.

## Technical note

The site uses the browser build of [Mammoth.js](https://github.com/mwilliamson/mammoth.js) from a public CDN.
