# Labor Tax Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a no-login web tool that accepts labor fee detail rows, calculates tax fields using the confirmed rules, previews results, and exports an Excel workbook with two sheets.

**Architecture:** Vue 3 provides one page for upload/manual entry/result preview. FastAPI exposes template, calculate-upload, calculate-manual, and export endpoints. A pure Decimal-based calculation service powers both JSON preview and Excel export.

**Tech Stack:** Vue 3 + Vite, Python 3.10+, FastAPI, openpyxl, pydantic, pytest.

## Global Constraints

- No login and no database.
- Input supports Excel upload and manual/pasted rows.
- Cumulative key is `year + month + id_no`.
- Calculation uses high precision internally and rounds display/export to 2 decimals.
- Export workbook has Sheet 1 strictly following the original ledger structure and Sheet 2 as a clearer business ledger.
- Export calculation cells contain final values only, not Excel formulas.

---

### Task 1: Backend calculation engine

**Files:**
- Create: `backend/app/schemas/labor.py`
- Create: `backend/app/services/tax_calculator.py`
- Test: `backend/tests/test_tax_calculator.py`

**Interfaces:**
- Produces: `LaborInputRow`, `LaborCalculatedRow`, `calculate_rows(rows: list[LaborInputRow]) -> list[LaborCalculatedRow]`.

- [x] Implement pydantic models and Decimal helpers.
- [x] Implement gross reverse calculation, individual income tax calculation, VAT, surcharge, invoice, payment, and check fields.
- [x] Test cumulative calculation across same person/month and reset across month/year.

### Task 2: Excel reader and writer

**Files:**
- Create: `backend/app/services/excel_reader.py`
- Create: `backend/app/services/excel_writer.py`
- Test: `backend/tests/test_excel_io.py`

**Interfaces:**
- Consumes: `LaborInputRow`, `LaborCalculatedRow`, `calculate_rows`.
- Produces: `read_labor_rows(file_obj) -> list[LaborInputRow]`, `build_template_workbook() -> bytes`, `build_result_workbook(rows) -> bytes`.

- [x] Parse Chinese and English field aliases from uploaded Excel.
- [x] Generate input template.
- [x] Generate result workbook with Sheet 1 original-style ledger and Sheet 2 clear ledger.

### Task 3: FastAPI endpoints

**Files:**
- Create: `backend/app/main.py`
- Create: `backend/app/api/labor_tax.py`
- Create: `backend/requirements.txt`

**Interfaces:**
- Consumes: Excel reader/writer and calculator services.
- Produces: `/api/template`, `/api/calculate/upload`, `/api/calculate/manual`, `/api/export`, `/api/health`.

- [x] Wire endpoints and JSON/file responses.
- [x] Return validation errors in a user-readable shape.

### Task 4: Vue 3 frontend

**Files:**
- Create: `frontend/package.json`
- Create: `frontend/index.html`
- Create: `frontend/vite.config.js`
- Create: `frontend/src/main.js`
- Create: `frontend/src/api/laborTaxApi.js`
- Create: `frontend/src/views/LaborTaxTool.vue`
- Create: `frontend/src/components/UploadPanel.vue`
- Create: `frontend/src/components/ManualTable.vue`
- Create: `frontend/src/components/ResultTable.vue`
- Create: `frontend/src/App.vue`

**Interfaces:**
- Consumes backend API endpoints.
- Produces one-page no-login UI.

- [x] Build upload/manual tabs, paste-friendly manual table, preview table, template/download/export buttons.

### Task 5: Documentation and packaging

**Files:**
- Create: `README.md`
- Create: `.gitignore`

**Interfaces:**
- Produces: runnable instructions and final zip package.

- [x] Document startup commands, input template columns, export behavior, and calculation rules.
