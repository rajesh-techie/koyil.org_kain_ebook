# Web Scraper to DOCX - Task Tracker

**Project:** URL-to-DOCX Web Scraper with Chapter Navigation  
**Standard:** Based on industry best practices (CIS/NIST for secure scripting)  
**Date Started:** 2026-04-02  
**Status Legend:** `[ ]` Not Started | `[~]` In Progress | `[x]` Done | `[-]` Skipped/Deferred

---

## TASK 0 — Technology Recommendation & Approval
- [x] **T0.1** — Present tool/language recommendation (Python vs alternatives)  
  - Recommended: Python 3.13.3 + Playwright + BeautifulSoup4 + python-docx  
  - **Approved by user 2026-04-02**

---

## TASK 1 — Project Setup & Environment
- [x] **T1.1** — Define final folder/file structure for the project — `src/scraper.py` layout created  
- [x] **T1.2** — Created `requirements.txt` with playwright, beautifulsoup4, lxml, python-docx  
- [x] **T1.3** — Created `README.md` with setup, usage, structure, and troubleshooting  
- [x] **T1.4** — Python 3.13.3 confirmed ✓
- [x] **T1.5** — Created `config.json` with all configurable keywords:
  - `start_markers`, `stop_markers`, `toc_selector`, `max_chapters`, `browser` settings
  - *(Requirement added 2026-04-02: all keywords must be externally configurable)*

---

## TASK 2 — Core Scraper: Input & Browser Launch
- [x] **T2.1** — CLI `--url` argument with http/https validation and SSRF guard  
- [x] **T2.2** — CLI `--output` argument with `.docx` extension check and writable-path validation  
- [x] **T2.3** — Launched Chromium headed browser via Playwright (headless controlled by config.json)  
- [x] **T2.4** — Navigated to URL with `wait_until="networkidle"` and logged page title — smoke tested on koyil.org ✓

---

## TASK 3 — Content Extraction Logic
- [x] **T3.1** — Reader Mode simulated: strips nav/footer/script tags, prefers `<article>`/`<main>`, deduplicates consecutive lines  
- [x] **T3.2** — Start marker: finds first line matching any `config.start_markers`, begins from one line before (or line 0) — tested: `varavaramunayE nama:` detected ✓  
- [x] **T3.3** — Stop marker: finds first line after start matching any `config.stop_markers`, stops one line before — tested: `ramanuja dasan` at line 108, stopped at 107 ✓  
- [x] **T3.4** — Extracted 107 clean lines from `https://divyaprabandham.koyil.org/index.php/2020/11/thiruviruththam/` ✓  
- [x] **T3.5** — Edge cases handled: marker not found (logs warning, returns empty or scrapes to end), empty page, timeout all handled ✓

---

## TASK 4 — Chapter Detection & Table of Contents
- [x] **T4.1** — Detect chapter links using `config.toc_selector` (`li > a[href]`), scoped to content area only  
  - Resolves relative → absolute URLs via `urljoin(base_url, href)`  
  - Same-domain filter: skips external/social links regardless of title text  
  - Deduplicates by normalised URL (fragment stripped)  
  - Title = `a.get_text()` — works for any chapter naming convention  
  - Capped at `config.max_chapters` (security guard)  
- [x] **T4.2** — Each detected link treated as a Chapter (Heading 2 in document)  
- [x] **T4.3** — TOC list built from chapter titles in order of appearance — **102 chapters detected on test page** ✓  
- [x] **T4.4** — Chapter iteration loop: opens each URL, applies start/stop marker rules, appends to chapter_contents list  
- [x] **T4.5** — Loop continues until all chapter links processed (or timeout/skip on error)

---

## TASK 5 — DOCX Generation
- [x] **T5.1** — `Document()` initialized via python-docx  
- [x] **T5.2** — Page 1: Heading 1 "Table of Contents" + Word TOC field (auto-updates page numbers in MS Word with F9) + plain-text fallback TOC list  
- [x] **T5.3** — Page 2: "Introduction" Heading 1 + main page scraped content  
- [x] **T5.4** — Pages 3+: Each chapter as Heading 2 + scraped body text; fallback message if no content  
- [x] **T5.5** — Consistent styling via Word built-in styles (Normal, Heading 1, Heading 2); paragraph indentation on TOC list  
- [x] **T5.6** — Saved to validated output filepath — **smoke tested: test_output.docx created** ✓

---

## TASK 6 — Error Handling & Logging
- [ ] **T6.1** — Add structured logging (Python `logging` module) for all major steps  
- [ ] **T6.2** — Handle network errors, timeouts, and missing pages gracefully  
- [ ] **T6.3** — Validate output path is writable before starting long scrape  
- [ ] **T6.4** — Log skipped chapters (if marker not found or page unreachable)

---

## TASK 7 — Security Hardening (CIS/NIST Alignment)
- [ ] **T7.1** — Validate and sanitize the input URL (prevent SSRF-style issues)  
- [ ] **T7.2** — Validate output file path (prevent directory traversal)  
- [ ] **T7.3** — Do not execute or eval any scraped content  
- [ ] **T7.4** — Cap maximum number of chapter links to process (prevent infinite loops)  
- [ ] **T7.5** — Run browser with restricted permissions (no file system access from browser context)

---

## TASK 8 — Testing
- [ ] **T8.1** — Test with the target URL provided by user  
- [ ] **T8.2** — Test start/stop marker detection with multiple page variations  
- [ ] **T8.3** — Test TOC and chapter scraping with multi-chapter pages  
- [ ] **T8.4** — Test output .docx opens correctly in Microsoft Word  
- [ ] **T8.5** — Test error handling paths (bad URL, missing markers, bad output path)

---

## TASK 9 — Final Packaging
- [ ] **T9.1** — Clean up code, add inline comments  
- [ ] **T9.2** — Finalize `README.md` with full usage examples  
- [ ] **T9.3** — Optional: Package as a single-file executable using `PyInstaller`

---

## Notes & Decisions Log
| Date | Decision | Reason |
|------|----------|--------|
| 2026-04-02 | Tasks.md created | Initial planning |
| 2026-04-02 | T0.1 approved | User approved Python + Playwright stack |
| 2026-04-02 | Task 1 complete | requirements.txt, README.md, src/ folder, scraper.py placeholder created |
| 2026-04-02 | Requirement refined | All scraping keywords (start, stop, TOC selector) must be externally configurable via config.json — T1.5 added |
| 2026-04-02 | Task 2 complete | scraper.py: CLI args, URL/path validation, Playwright Chromium launch, page navigation — smoke tested on koyil.org |
| 2026-04-02 | Task 3 complete | Content extraction: parse_lines (Reader Mode sim), start/stop marker detection — live tested on divyaprabandham.koyil.org, 107 lines extracted correctly |
| 2026-04-02 | Task 4 complete | Chapter detection: scoped to content area, same-domain filter, relative URL resolution, deduplicated — 102/102 chapters detected on test page |
| 2026-04-02 | Task 5 complete | DOCX built: TOC page (Word field + plain list), Introduction page, chapter sections (Heading 2) — smoke tested with 3 chapters, test_output.docx saved |

