# PRD: Legal Contract Translator Hybrid

## 1. Product Overview

Legal Contract Translator Hybrid is a desktop tool for translating Chinese legal contracts into English while preserving the original Word document structure and key formatting. The product targets capital-markets and corporate legal documents where translation quality, context consistency, and Word formatting fidelity are all required.

The tool combines large-language-model translation with deterministic `.docx` processing. It uses Markdown to preserve document hierarchy and context, Word run analysis to capture inline formatting, and a post-translation review pipeline to detect untranslated Chinese text and machine-readable artifacts.

## 2. Background

Manual legal translation of long Word contracts is slow and error-prone, especially when documents contain tables, headings, placeholders, bold text, underline, highlights, page headers, page footers, and complex numbering. Direct paragraph-by-paragraph translation loses context, while direct run-by-run translation preserves formatting but produces fragmented translation quality.

The selected implementation uses a hybrid method:

- Markdown blocks provide clause hierarchy, paragraph context, indentation cues, and large-chunk translation context.
- Word run structures provide deterministic extraction of formatting spans such as bold, underline, highlight, and run-level styling.
- The model translates large Markdown chunks and separately maps source formatted spans to corresponding English phrases.
- The program writes translated text back into the original Word document and reapplies mapped formatting.
- A final review step checks for residual Chinese/CJK text and leaked intermediate machine markers.

## 3. Goals

- Translate Chinese `.docx` legal contracts into legal English with strong contextual consistency.
- Preserve Word document structure and key visible formatting.
- Provide a workflow usable by non-technical users: input API Key, add Word files, choose output folder, start translation.
- Support both Windows and macOS.
- Support multi-file processing with independent progress display and cancellation.
- Export a final translated Word document and process files for audit and debugging.
- Avoid uploading local API keys or generated outputs to source control.

## 4. Non-Goals

- The product is not a substitute for qualified legal review.
- The product does not guarantee perfect preservation of every low-level Word XML object.
- The product does not currently target scanned PDFs or image OCR workflows.
- The product does not implement a full CAT/TM production system.
- The product does not require cloud storage or a hosted backend.

## 5. Target Users

- Lawyers and legal translators handling Chinese-to-English contract translation.
- Capital-markets or corporate legal teams working with investment agreements, shareholder agreements, certificates, schedules, and appendices.
- Non-technical users who can provide an API Key and operate a simple desktop interface.
- Team members on Windows or macOS who need a local translation tool.

## 6. Core User Workflow

1. User launches the tool through the platform-specific startup script.
2. User selects API provider and enters API Key, Base URL, and model.
3. User selects English and Chinese font preferences.
4. User adds one or more `.docx` files.
5. User selects output directory.
6. User starts translation.
7. Each file displays independent translation progress and review progress.
8. User may cancel all tasks or cancel a single file and export current progress.
9. The tool outputs translated Word files and process files.
10. User reviews the exported Word document and optional checklist.

## 7. Functional Requirements

### P0 Requirements

| ID | Requirement | Description | Acceptance Criteria |
|---|---|---|---|
| P0-01 | API configuration | Support OpenAI-compatible and DeepSeek-compatible API settings. | User can input API Key, Base URL, and model; DeepSeek defaults to `https://api.deepseek.com` and `deepseek-v4-flash`. |
| P0-02 | Word input | Support `.docx` files. | User can add at least one `.docx`; invalid files are rejected with a clear message. |
| P0-03 | Hybrid translation | Translate using Markdown context plus run-format mapping. | Long legal clauses are translated with contextual coherence while formatted spans are reapplied to English text. |
| P0-04 | Format preservation | Preserve major visible Word formatting. | Output preserves paragraph structure, tables, bold, underline, highlight, headers, footers, placeholders, numbering, and key fonts where technically supported. |
| P0-05 | Final review | Detect residual Chinese/CJK text and leaked machine artifacts. | Final review scans translated output for Chinese/CJK content and markers such as `<!-- META:... -->`, `<!-- BLOCK:... -->`, JSON snippets, and fenced code blocks. |
| P0-06 | Export | Generate translated Word document. | Output folder contains final `.docx` for each source file. |
| P0-07 | Process files | Export audit files separately. | Markdown source, translated Markdown, checklist, and JSON details are saved under a `过程文件` folder. |
| P0-08 | Cancellation | Allow cancellation with current progress export. | User can cancel all tasks or one task; completed progress is exported without waiting for unrelated future work. |
| P0-09 | Windows launcher | Provide Windows startup script. | `一键启动.bat` creates/checks environment and launches the Windows interface. |
| P0-10 | macOS launcher | Provide macOS startup script. | `一键启动.command` creates/checks environment and launches the macOS interface. |

### P1 Requirements

| ID | Requirement | Description | Acceptance Criteria |
|---|---|---|---|
| P1-01 | Multi-file processing | Support adding multiple files and translating them in parallel. | Each file has its own progress row, status, and cancellation button. |
| P1-02 | Independent progress display | Separate translation and review progress. | UI shows one progress bar for translation and one progress bar for review per file. |
| P1-03 | File list management | Allow adding, deleting selected files, and clearing the list. | User can remove mistakenly added files before starting. |
| P1-04 | Font preferences | Allow selecting English and Chinese fonts. | English font supports Times New Roman and Calibri; Chinese font supports common Chinese legal-document fonts; digits use Times New Roman. |
| P1-05 | API validation | Provide API Key test for macOS interface. | User can test API credentials before full translation; authentication failures are shown clearly. |
| P1-06 | macOS-specific interface | Avoid macOS-incompatible drag-and-drop dependencies. | macOS interface works without `tkinterdnd2`; files are selected by button. |
| P1-07 | Scrollable UI | Support scrolling on small screens. | Long forms and progress sections can be scrolled with mouse wheel or trackpad. |

### P2 Requirements

| ID | Requirement | Description | Acceptance Criteria |
|---|---|---|---|
| P2-01 | Glossary support | Improve consistency of legal terminology and defined terms. | Prompt includes terminology guidance and translation memory generated from source context. |
| P2-02 | Company-name handling | Avoid repeated repair loops caused by bilingual company-name formatting. | Final policy avoids forcing English-name-plus-Chinese-name output unless explicitly required. |
| P2-03 | Better audit reporting | Improve checklist readability for non-technical users. | Checklist explains remaining risks in human-readable language. |
| P2-04 | Packaging | Provide colleague-ready ZIP packages. | Packages contain only required runtime files and exclude `.venv`, local config, saved keys, and output files. |

## 8. Translation Pipeline

### 8.1 Document Parsing

- Load `.docx` with `python-docx`.
- Iterate through body, tables, headers, and footers.
- Extract paragraphs containing Chinese/CJK text.
- Convert relevant paragraphs into Markdown blocks with IDs and metadata.
- Extract run-level formatting spans, including bold, underline, highlight, and other visible styles.

### 8.2 Context Translation

- Batch Markdown blocks into large chunks.
- Send each chunk to the selected API model.
- Require legal English output and prohibit untranslated Chinese/CJK text.
- Preserve placeholders, clause references, numbering, dates, amounts, and defined terms.

### 8.3 Format Mapping

- For each block, provide original formatted spans and translated English text.
- Ask the model to locate the corresponding English phrase for each formatted Chinese span.
- Programmatically apply formatting to the mapped English phrase.
- Record mapping status in checklist.

### 8.4 Review and Repair

- Run automated checks for missing translations and Chinese/CJK leftovers.
- Run final LLM audit for untranslated legal content.
- Run final Word sweep over visible document text.
- Detect and remove leaked machine artifacts, including Markdown comments, block markers, metadata comments, JSON snippets, and fenced code blocks.
- Export with checklist notes if any residual risks remain.

## 9. Output Specification

For each source file, the tool should generate:

- Final translated Word document: `*_复合方法英文翻译.docx`
- Source block Markdown: `*_source_blocks.md`
- Translated block Markdown: `*_translated_blocks.md`
- Checklist: `*_复合方法checklist.md`
- JSON details: `*_复合方法明细.json`

Process files should be stored under a dedicated `过程文件` folder. Users may delete process files after confirming the final Word output.

## 10. UI Requirements

### Windows Interface

- Supports API settings, file selection, output selection, font options, multi-file progress, single-file cancellation, and all-task cancellation.
- Supports drag-and-drop where available.
- Uses the shared translation core.

### macOS Interface

- Uses a dedicated macOS GUI to avoid cross-platform Tk issues.
- Does not depend on drag-and-drop plugins.
- Uses direct entry widgets for API Key, Base URL, model, and output directory.
- Provides a `Test API Key` action.
- Provides full-page scrolling and per-file progress display.

## 11. Error Handling Requirements

- Authentication failures should be displayed as API credential errors, not generic stack traces.
- Missing API Key, missing model, missing files, invalid files, and inaccessible output folders should be handled before translation starts.
- If output file is locked by Word/WPS/OneDrive, the tool should attempt safe saving or report a clear failure.
- Cancellation should not corrupt existing output files.
- API failures should be logged per file without terminating unrelated files.

## 12. Security and Privacy

- API Key is entered locally by the user.
- If "remember API Key" is enabled, the key is stored only in local config files.
- Local config files and generated outputs must be ignored by Git.
- The tool should not upload documents anywhere except to the selected LLM API endpoint.
- Public repository must not contain API keys, user documents, generated output files, or local virtual environments.

## 13. Acceptance Criteria

- A non-technical user can start the tool on Windows or macOS.
- User can add multiple `.docx` files and start translation.
- Each file displays translation progress and review progress.
- User can cancel one file or all files and still export current progress.
- Output Word document preserves major document structure and visible formatting.
- Final review checks both Chinese/CJK leftovers and machine-readable artifacts.
- Process files are stored separately from final Word output.
- Repository can be made public without exposing local API keys or generated client documents.

## 14. Version Milestones

| Version | Key Outcome |
|---|---|
| v1.13 | Strengthened final English-only sweep. |
| v1.18 | Added multi-file parallel progress and file removal. |
| v1.19 | Added machine-artifact detection and cleanup. |
| v1.20 | Removed obsolete alternative schemes and kept only the hybrid method. |
| v1.22 | Added macOS launcher. |
| v1.26 | Added dedicated macOS interface. |
| v1.28 | Improved macOS touchpad scrolling. |

## 15. Open Questions

- Whether to add a formal terminology database or user-editable glossary.
- Whether to support bilingual review exports.
- Whether to support PDF/OCR input in a future version.
- Whether to convert the desktop tool into a packaged executable for easier distribution.
- Whether to add automated visual regression checks for Word layout fidelity.
