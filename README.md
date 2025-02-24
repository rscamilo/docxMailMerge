# docx-mail-merge

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

`docx-mail-merge` is a lightweight, dependency-free (except for `jszip`, which is essential) JavaScript library for performing mail merge operations directly in the browser (or in Node.js) using DOCX templates and JSON data. It generates new DOCX files by replacing placeholders in a template with text, images, and tables. This eliminates the need for server-side processing for mail merge, improving performance and reducing server load.

## Features

*   **Client-Side Mail Merge:** Performs the entire mail merge process in the browser, eliminating server requests for each document generation.
*   **JSON Data Input:** Uses JSON objects to provide data for the mail merge, making it easy to integrate with various data sources.
*   **DOCX Template Support:**  Works with standard DOCX files as templates.  You can create your templates using Microsoft Word or any other compatible editor.
*   **Text Replacement:** Replaces simple text placeholders (e.g., `{{NAME}}`, `{{ADDRESS}}`) with corresponding values from the JSON data. Supports multiline text with preserved formatting.
*   **Image Replacement:** Replaces placeholders with images provided as base64-encoded data URIs.  Allows specifying image width and height.
*   **Table Generation:** Replaces placeholders with dynamically generated tables from array data.  Creates properly formatted Word tables with borders and cell styling.
*   **Base64 Input and Output:** Accepts the DOCX template as a base64-encoded data URI and returns the generated document as a base64-encoded data URI. This is ideal for web applications where you might want to fetch the template from a server or database.
*   **Asynchronous Operation:** Uses `async/await` for non-blocking operations, ensuring a smooth user experience.
*   **Error Handling:** Includes error handling to catch invalid template formats and other potential issues.
*   **Node.js and Browser Compatible:**  Can be used in both Node.js and browser environments (where `require('jszip')` is replaced with a suitable import in the browser context, see Installation section).

## Installation

```bash
npm install jszip --save
