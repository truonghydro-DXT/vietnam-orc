import base64
import json
import re
import socket
import threading
import time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from export_ocr import make_docx, make_json, make_xlsx


APP_DIR = Path(__file__).resolve().parent
HOST = "0.0.0.0"
PORT = 8000
UPLOAD_DIR = APP_DIR / "runs" / "web" / "uploads"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
PDF_EXTENSIONS = {".pdf"}
OCR_UPSCALE_FACTOR = 1.5
OCR_MAX_SIDE = 2600
OCR_DET_LIMIT_SIDE = 1600
DOCUMENT_API_ROUTES = {
    "/api/cccd-front": "cccd_front",
    "/api/cccd-back": "cccd_back",
    "/api/land-certificate": "land_certificate_1",
    "/api/land-certificate-2": "land_certificate_2",
    "/api/land-certificate-3": "land_certificate_3",
}


INDEX_HTML = """<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Vietnamese OCR</title>
  <style>
    :root {
      color-scheme: light;
      font-family: Arial, Helvetica, sans-serif;
      --border: #d7dce2;
      --primary: #2563eb;
      --muted: #667085;
      --panel: #ffffff;
      --bg: #f3f6fa;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      background: var(--bg);
      color: #101828;
    }

    header {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 14px 18px;
      background: var(--panel);
      border-bottom: 1px solid var(--border);
      position: sticky;
      top: 0;
      z-index: 2;
    }

    h1 {
      font-size: 18px;
      margin: 0 12px 0 0;
      white-space: nowrap;
    }

    button, select, input::file-selector-button {
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #fff;
      color: #101828;
      cursor: pointer;
      font-size: 14px;
      padding: 9px 12px;
    }

    button.primary {
      background: var(--primary);
      border-color: var(--primary);
      color: #fff;
      font-weight: 600;
    }

    button:disabled {
      cursor: not-allowed;
      opacity: 0.55;
    }

    select {
      min-width: 110px;
    }

    #sourceSelect {
      min-width: 180px;
    }

    #fileInput {
      display: none;
    }

    #status {
      color: var(--muted);
      flex: 1;
      min-width: 160px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    main {
      display: grid;
      grid-template-columns: 4fr 7fr;
      gap: 12px;
      height: calc(100vh - 67px);
      padding: 12px;
    }

    section {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 12px;
      min-width: 0;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }

    section h2 {
      border-bottom: 1px solid var(--border);
      font-size: 15px;
      margin: 0;
      padding: 10px 12px;
    }

    .section-head {
      align-items: center;
      border-bottom: 1px solid var(--border);
      display: flex;
      gap: 8px;
      justify-content: space-between;
      padding: 8px 12px;
    }

    .section-head h2 {
      border: 0;
      padding: 0;
    }

    .page-nav {
      align-items: center;
      display: flex;
      gap: 6px;
    }

    .page-nav button {
      min-width: 34px;
      padding: 6px 8px;
    }

    .page-nav span {
      color: var(--muted);
      font-size: 13px;
      min-width: 88px;
      text-align: center;
      white-space: nowrap;
    }

    .image-wrap {
      align-items: center;
      background: #f8fafc;
      display: flex;
      flex: 1;
      justify-content: center;
      min-height: 0;
      overflow: hidden;
      padding: 12px;
    }

    .image-stage {
      align-items: center;
      display: flex;
      height: 100%;
      justify-content: center;
      max-height: 100%;
      max-width: 100%;
      position: relative;
      width: 100%;
    }

    #preview {
      display: none;
      height: 100%;
      width: 100%;
      object-fit: contain;
    }

    #overlay {
      display: none;
      height: 100%;
      left: 0;
      pointer-events: none;
      position: absolute;
      top: 0;
      width: 100%;
    }

    .ocr-box {
      fill: rgba(37, 99, 235, 0.08);
      stroke: #2563eb;
      stroke-width: 2;
    }

    .table-line {
      stroke: #dc2626;
      stroke-linecap: round;
      stroke-width: 3;
    }

    .placeholder {
      color: var(--muted);
      text-align: center;
    }

    .formatted-wrap {
      align-items: flex-start;
      display: flex;
      flex-direction: column;
      flex: 1;
      min-height: 260px;
      overflow-x: hidden;
      overflow-y: auto;
      padding: 12px;
    }

    .formatted-wrap h3 {
      font-size: 14px;
      margin: 0 0 8px;
      width: 100%;
    }

    #formattedPlaceholder {
      color: var(--muted);
      font-size: 14px;
    }

    #formattedResult {
      background: transparent;
      border: 0;
      box-shadow: none;
      display: none;
      overflow: visible;
      position: relative;
      transform-origin: top left;
    }

    .formatted-text {
      color: #111827;
      font-family: "Times New Roman", Times, serif;
      line-height: 1.15;
      overflow: hidden;
      position: absolute;
      white-space: pre-wrap;
      word-break: break-word;
    }

    .formatted-title-red {
      color: #dc2626;
      font-family: Arial, Helvetica, sans-serif;
      font-weight: 800;
      letter-spacing: 0.2px;
      text-align: center;
      text-transform: uppercase;
    }

    .formatted-heading {
      font-family: "Times New Roman", Times, serif;
      font-weight: 700;
      text-align: center;
    }

    .formatted-section {
      font-family: "Times New Roman", Times, serif;
      font-weight: 700;
    }

    .formatted-table-text {
      font-family: "Times New Roman", Times, serif;
      font-weight: 400;
      text-align: center;
    }

    .formatted-grid-line {
      background: #111827;
      position: absolute;
      transform-origin: top left;
    }

    .formatted-table-block {
      border-left: 1px solid #111827;
      border-top: 1px solid #111827;
      display: grid;
      overflow: hidden;
      position: absolute;
    }

    .formatted-table-cell {
      border-bottom: 1px solid #111827;
      border-right: 1px solid #111827;
      overflow: hidden;
      position: relative;
    }

    .spacer {
      flex: 1;
    }

    @media (max-width: 900px) {
      header {
        align-items: stretch;
        flex-wrap: wrap;
      }

      main {
        grid-template-columns: 1fr;
        height: auto;
        min-height: calc(100vh - 67px);
      }

      section {
        min-height: 420px;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>Vietnamese OCR</h1>
    <select id="sourceSelect">
      <option value="" selected>Chọn nguồn...</option>
      <option value="file">1. Chọn ảnh/pdf</option>
      <option value="printer">2. Kết nối máy in</option>
      <option value="scanner">3. Kết nối máy scan</option>
    </select>
    <input id="fileInput" type="file" accept="image/*,.pdf,application/pdf" multiple>
    <button id="ocrButton" class="primary" disabled>OCR</button>
    <button id="exportDocx" disabled>Export Word</button>
    <button id="exportXlsx" disabled>Export Excel</button>
    <button id="exportJson" disabled>Export JSON</button>
    <select id="jsonType">
      <option value="cccd_front" selected>Mặt trước CCCD</option>
      <option value="cccd_back">Mặt sau CCCD</option>
      <optgroup label="Giấy CNSDĐ">
        <option value="land_certificate_1">Mẫu số 1</option>
        <option value="land_certificate_2">Mẫu số 2 (nhiều trang / PDF)</option>
        <option value="land_certificate_3">Mẫu số 3</option>
      </optgroup>
    </select>
    <select id="engine">
      <option value="vietocr" selected>vietocr</option>
      <option value="paddle">paddle</option>
    </select>
    <span id="status">Đang nạp VietOCR...</span>
  </header>

  <main>
    <section>
      <div class="section-head">
        <h2>Ảnh gốc</h2>
        <div class="page-nav">
          <button type="button" id="prevPage" disabled>‹</button>
          <span id="pageLabel">Trang 0/0</span>
          <button type="button" id="nextPage" disabled>›</button>
        </div>
      </div>
      <div class="image-wrap">
        <div class="image-stage">
          <img id="preview" alt="Ảnh gốc">
          <svg id="overlay"></svg>
        </div>
        <div id="placeholder" class="placeholder"></div>
      </div>
    </section>

    <section>
      <div class="section-head">
        <h2>Kết quả OCR</h2>
        <div class="page-nav">
          <span id="resultPageLabel">Chưa OCR</span>
        </div>
      </div>
      <div class="formatted-wrap">
        <div id="formattedPlaceholder"></div>
        <div id="formattedResult"></div>
      </div>
    </section>
  </main>

  <script>
    const sourceSelect = document.getElementById("sourceSelect");
    const fileInput = document.getElementById("fileInput");
    const ocrButton = document.getElementById("ocrButton");
    const exportDocx = document.getElementById("exportDocx");
    const exportXlsx = document.getElementById("exportXlsx");
    const exportJson = document.getElementById("exportJson");
    const jsonType = document.getElementById("jsonType");
    const engine = document.getElementById("engine");
    const statusEl = document.getElementById("status");
    const preview = document.getElementById("preview");
    const overlay = document.getElementById("overlay");
    const placeholder = document.getElementById("placeholder");
    const formattedResult = document.getElementById("formattedResult");
    const formattedPlaceholder = document.getElementById("formattedPlaceholder");
    const prevPage = document.getElementById("prevPage");
    const nextPage = document.getElementById("nextPage");
    const pageLabel = document.getElementById("pageLabel");
    const resultPageLabel = document.getElementById("resultPageLabel");

    let selectedFile = null;
    let selectedFiles = [];
    let currentPageIndex = 0;
    let selectedDataUrl = "";
    let latestText = "";
    let latestBoxes = [];
    let latestTableLines = [];
    let latestTableRows = [];
    let latestFormattedData = null;

    function isMau2MultiPage() {
      return jsonType.value === "land_certificate_2";
    }

    function mau2PagesForJson() {
      // Mẫu 2: gộp OCR mọi trang đã nhận diện (PDF nhiều trang hoặc 2 ảnh).
      return selectedFiles.filter((page) => String(page?.ocrData?.text || "").trim());
    }

    function currentPage() {
      return selectedFiles[currentPageIndex] || null;
    }

    async function waitForPreview() {
      if (preview.complete && preview.naturalWidth > 0) {
        return;
      }
      if (typeof preview.decode === "function") {
        try {
          await preview.decode();
          return;
        } catch (_) {
          /* fall through to load/error events */
        }
      }
      await new Promise((resolve) => {
        preview.addEventListener("load", resolve, { once: true });
        preview.addEventListener("error", resolve, { once: true });
      });
    }

    function updatePageNav() {
      const total = selectedFiles.length;
      const index = total ? currentPageIndex + 1 : 0;
      pageLabel.textContent = `Trang ${index}/${total}`;
      prevPage.disabled = currentPageIndex <= 0;
      nextPage.disabled = currentPageIndex >= total - 1 || total === 0;
      const page = currentPage();
      if (!page) {
        resultPageLabel.textContent = "Chưa OCR";
      } else if (page.ocrData) {
        resultPageLabel.textContent = `Kết quả trang ${index}/${total}`;
      } else {
        resultPageLabel.textContent = `Chưa OCR trang ${index}`;
      }
    }

    function syncExportFromPages() {
      if (isMau2MultiPage()) {
        const mau2Pages = mau2PagesForJson();
        if (mau2Pages.length >= 2) {
          latestText = mau2Pages.map((page) => page.ocrData.text || "").filter(Boolean).join("\\n");
          latestTableRows = mau2Pages.flatMap((page) => page.ocrData.table_rows || []);
          latestFormattedData = currentPage()?.ocrData || mau2Pages[0].ocrData || null;
        } else {
          const page = currentPage();
          latestText = page?.ocrData?.text || "";
          latestTableRows = page?.ocrData?.table_rows || [];
          latestFormattedData = page?.ocrData || null;
        }
      } else {
        const page = currentPage();
        latestText = page?.ocrData?.text || "";
        latestTableRows = page?.ocrData?.table_rows || [];
        latestFormattedData = page?.ocrData || null;
      }
      updateExportButtons();
    }

    async function showPage(index) {
      if (!selectedFiles.length) {
        return;
      }
      currentPageIndex = Math.max(0, Math.min(index, selectedFiles.length - 1));
      const pageIndex = currentPageIndex;
      const page = selectedFiles[pageIndex];
      selectedFile = page.file;
      selectedDataUrl = page.displayImage || page.image;
      preview.src = selectedDataUrl;
      preview.style.display = "block";
      placeholder.style.display = "none";
      ocrButton.disabled = false;
      updatePageNav();

      clearOverlay();
      clearFormattedResult();
      await waitForPreview();
      if (currentPageIndex !== pageIndex) {
        return;
      }
      if (page.ocrData) {
        const data = page.ocrData;
        drawOverlay(data.boxes || [], data.table_lines || []);
        renderFormattedResult(data);
      }
      syncExportFromPages();
      setStatus(`Đang xem: ${page.name} (trang ${pageIndex + 1}/${selectedFiles.length})`);
    }

    jsonType.addEventListener("change", () => {
      if (isMau2MultiPage()) {
        setStatus("Mẫu số 2: chọn ≥ 2 ảnh, hoặc 1 PDF (≥ 2 trang), rồi OCR.");
      }
      syncExportFromPages();
    });

    prevPage.addEventListener("click", () => {
      showPage(currentPageIndex - 1);
    });
    nextPage.addEventListener("click", () => {
      showPage(currentPageIndex + 1);
    });

    function updateExportButtons() {
      const hasPageText = selectedFiles.some((page) => String(page?.ocrData?.text || "").trim().length > 0);
      const hasText = hasPageText || latestText.trim().length > 0;
      exportDocx.disabled = !hasText;
      exportXlsx.disabled = !hasText;
      exportJson.disabled = !hasText;
    }

    function collectExportPages() {
      const pages = [];
      selectedFiles.forEach((page, index) => {
        const text = String(page?.ocrData?.text || "").trim();
        if (!text) {
          return;
        }
        const data = page.ocrData;
        pages.push({
          title: `Trang ${index + 1}: ${page.name}`,
          sheet_name: `Trang ${index + 1}`,
          text: data.text || "",
          table_rows: data.table_rows || [],
          formatted_data: data
        });
      });
      if (pages.length) {
        return pages;
      }
      if (latestText.trim()) {
        return [{
          title: "Trang 1",
          sheet_name: "Trang 1",
          text: latestText,
          table_rows: latestTableRows,
          formatted_data: latestFormattedData || {}
        }];
      }
      return [];
    }

    function setStatus(message) {
      statusEl.textContent = message;
    }

    function readFileAsDataUrl(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(file);
      });
    }

    function clearOverlay() {
      latestBoxes = [];
      latestTableLines = [];
      overlay.innerHTML = "";
      overlay.style.display = "none";
    }

    function clearResultData() {
      latestText = "";
      latestTableRows = [];
    }

    function clearFormattedResult() {
      latestFormattedData = null;
      formattedResult.innerHTML = "";
      formattedResult.style.display = "none";
      formattedPlaceholder.style.display = "block";
    }

    function normalizeBox(box) {
      if (!Array.isArray(box) || box.length === 0) {
        return null;
      }

      let points = box;
      if (box.length === 2 && Array.isArray(box[0]) && Array.isArray(box[1])) {
        points = [
          [box[0][0], box[0][1]],
          [box[1][0], box[0][1]],
          [box[1][0], box[1][1]],
          [box[0][0], box[1][1]]
        ];
      }

      const xs = points.map((point) => Number(point[0]));
      const ys = points.map((point) => Number(point[1]));
      if (xs.some(Number.isNaN) || ys.some(Number.isNaN)) {
        return null;
      }

      return {
        x: Math.min(...xs),
        y: Math.min(...ys),
        width: Math.max(...xs) - Math.min(...xs),
        height: Math.max(...ys) - Math.min(...ys)
      };
    }

    function getDisplayedImageRect() {
      const stageRect = preview.parentElement.getBoundingClientRect();
      const naturalWidth = preview.naturalWidth;
      const naturalHeight = preview.naturalHeight;
      if (!stageRect.width || !stageRect.height || !naturalWidth || !naturalHeight) {
        return null;
      }

      const scale = Math.min(stageRect.width / naturalWidth, stageRect.height / naturalHeight);
      const width = naturalWidth * scale;
      const height = naturalHeight * scale;
      return {
        left: (stageRect.width - width) / 2,
        top: (stageRect.height - height) / 2,
        width,
        height
      };
    }

    function drawOverlay(boxes, tableLines = latestTableLines) {
      latestBoxes = boxes || [];
      latestTableLines = tableLines || [];
      overlay.innerHTML = "";

      if (!preview.naturalWidth || !preview.naturalHeight || (!latestBoxes.length && !latestTableLines.length)) {
        overlay.style.display = "none";
        return;
      }

      const imageRect = getDisplayedImageRect();
      if (!imageRect) {
        overlay.style.display = "none";
        return;
      }

      overlay.style.left = `${imageRect.left}px`;
      overlay.style.top = `${imageRect.top}px`;
      overlay.style.width = `${imageRect.width}px`;
      overlay.style.height = `${imageRect.height}px`;
      overlay.setAttribute("viewBox", `0 0 ${preview.naturalWidth} ${preview.naturalHeight}`);
      for (const box of latestBoxes) {
        const rect = normalizeBox(box);
        if (!rect) {
          continue;
        }
        const element = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        element.setAttribute("class", "ocr-box");
        element.setAttribute("x", rect.x);
        element.setAttribute("y", rect.y);
        element.setAttribute("width", rect.width);
        element.setAttribute("height", rect.height);
        overlay.appendChild(element);
      }

      for (const line of latestTableLines) {
        if (!Array.isArray(line) || line.length < 4) {
          continue;
        }
        const element = document.createElementNS("http://www.w3.org/2000/svg", "line");
        element.setAttribute("class", "table-line");
        element.setAttribute("x1", line[0]);
        element.setAttribute("y1", line[1]);
        element.setAttribute("x2", line[2]);
        element.setAttribute("y2", line[3]);
        overlay.appendChild(element);
      }
      overlay.style.display = "block";
    }

    function renderFormattedResult(data) {
      latestFormattedData = data;
      formattedResult.innerHTML = "";

      const imageWidth = Number(data.image_width || preview.naturalWidth);
      const imageHeight = Number(data.image_height || preview.naturalHeight);
      const items = data.items || [];
      const layoutRows = data.layout_rows || [];
      const tableCells = data.table_cells || [];
      const tableGrid = data.table_grid || null;
      const lines = data.table_lines || [];

      if (!imageWidth || !imageHeight || (!layoutRows.length && !tableCells.length && !items.length && !lines.length && !tableGrid)) {
        formattedResult.style.display = "none";
        formattedPlaceholder.style.display = "block";
        return;
      }

      const containerWidth = formattedResult.parentElement.clientWidth - 24;
      const targetWidth = containerWidth * 0.9;
      const widthScale = targetWidth / imageWidth;
      const scale = Math.min(1, Math.max(0.1, widthScale));
      formattedResult.style.width = `${imageWidth * scale}px`;
      formattedResult.style.height = `${imageHeight * scale}px`;
      formattedResult.style.display = "block";
      formattedPlaceholder.style.display = "none";

      function classifyText(text) {
        const upper = String(text || "").toUpperCase();
        const letters = upper.replace(/[^A-ZÀ-ỴĐ]/g, "");
        const originalLetters = String(text || "").replace(/[^A-Za-zÀ-ỴĐà-ỵđ]/g, "");
        const upperRatio = originalLetters.length
          ? originalLetters.split("").filter((char) => char.toUpperCase() === char).length / originalLetters.length
          : 0;
        if (upperRatio > 0.65 && letters.length >= 8) {
          return "heading";
        }
        if (/^\\s*\\d+\\.\\s|^\\s*[A-ZĐ]\\.\\s/.test(upper)) {
          return "section";
        }
        return "normal";
      }

      function applyTextStyle(element, text, fontSize, options = {}) {
        const sourceStyle = options.style || {};
        const kind = sourceStyle.kind || options.kind || classifyText(text);
        element.classList.add("formatted-text");
        element.style.fontSize = `${fontSize}px`;
        element.style.lineHeight = `${Math.max(fontSize * 1.28, 8)}px`;
        element.style.color = sourceStyle.color || "#111827";

        if (options.table) {
          element.classList.add("formatted-table-text");
          if (sourceStyle.bold) {
            element.style.fontWeight = "700";
          }
          return;
        }

        if (sourceStyle.color && sourceStyle.color.toLowerCase() !== "#111827" && (sourceStyle.bold || kind === "heading")) {
          element.classList.add("formatted-title-red");
          element.style.fontSize = `${fontSize * 1.35}px`;
          element.style.lineHeight = `${fontSize * 1.55}px`;
          element.style.color = sourceStyle.color;
        } else if (kind === "heading") {
          element.classList.add("formatted-heading");
        } else if (kind === "section") {
          element.classList.add("formatted-section");
        }
        if (sourceStyle.bold) {
          element.style.fontWeight = "700";
        }
        if (sourceStyle.align === "center") {
          element.style.textAlign = "center";
        }
      }

      function createTableElement(tableTop) {
        const tableElement = document.createElement("div");
        tableElement.className = "formatted-table-block";
        tableElement.style.left = `${Number(tableGrid.x || 0) * scale}px`;
        tableElement.style.top = `${tableTop}px`;
        tableElement.style.width = `${Number(tableGrid.w || 0) * scale}px`;
        tableElement.style.height = `${Number(tableGrid.h || 0) * scale}px`;
        tableElement.style.gridTemplateColumns = tableGrid.columns.map((value) => `${Math.max(1, Number(value) * scale)}px`).join(" ");
        tableElement.style.gridTemplateRows = tableGrid.rows.map((value) => `${Math.max(1, Number(value) * scale)}px`).join(" ");

        const cellMap = new Map(tableCells.map((cell) => [`${cell.row}:${cell.col}`, cell]));
        for (let rowIndex = 0; rowIndex < tableGrid.rows.length; rowIndex += 1) {
          for (let colIndex = 0; colIndex < tableGrid.columns.length; colIndex += 1) {
            const gridCell = document.createElement("div");
            gridCell.className = "formatted-table-cell";
            const cell = cellMap.get(`${rowIndex}:${colIndex}`);
            if (cell) {
              const cellLines = (cell.lines && cell.lines.length ? cell.lines : [cell.text || ""])
                .map((line) => String(line || "").trim())
                .filter(Boolean);
              if (cellLines.length) {
                const cellWidth = Math.max(Number(tableGrid.columns[colIndex] || 1) * scale, 24);
                const cellHeight = Math.max(Number(tableGrid.rows[rowIndex] || 1) * scale, 12);
                const longestLine = cellLines.reduce((longest, line) => Math.max(longest, line.length), 1);
                const fontByHeight = (cellHeight * 0.82) / Math.max(1, cellLines.length * 1.25);
                const fontByWidth = cellWidth / Math.max(4, longestLine * 0.62);
                const fontSize = Math.max(4.5, Math.min(11, fontByHeight, fontByWidth));
                const lineHeight = Math.max(fontSize * 1.25, 6);
                const contentHeight = lineHeight * cellLines.length;
                const element = document.createElement("div");
                applyTextStyle(element, cellLines.join("\\n"), fontSize, {table: true, style: cell.style || {}});
                element.textContent = cellLines.join("\\n");
                element.style.left = "0px";
                element.style.top = `${Math.max(1, (cellHeight - contentHeight) / 2)}px`;
                element.style.width = `${cellWidth}px`;
                element.style.minHeight = `${contentHeight}px`;
                element.style.lineHeight = `${lineHeight}px`;
                element.style.whiteSpace = "pre-wrap";
                element.style.overflow = "visible";
                gridCell.appendChild(element);
              }
            }
            tableElement.appendChild(gridCell);
          }
        }

        return tableElement;
      }

      function finishFormattedHeight(contentBottom = 0) {
        const bottomPadding = Math.max(32, imageHeight * scale * 0.04);
        const targetHeight = Math.max(imageHeight * scale, contentBottom + bottomPadding);
        formattedResult.style.height = `${targetHeight}px`;
      }

      if (!tableGrid) {
        for (const line of lines) {
          if (!Array.isArray(line) || line.length < 4) {
            continue;
          }
          const x1 = Number(line[0]) * scale;
          const y1 = Number(line[1]) * scale;
          const x2 = Number(line[2]) * scale;
          const y2 = Number(line[3]) * scale;
          const element = document.createElement("div");
          element.className = "formatted-grid-line";
          element.style.left = `${Math.min(x1, x2)}px`;
          element.style.top = `${Math.min(y1, y2)}px`;
          if (Math.abs(y2 - y1) <= Math.abs(x2 - x1)) {
            element.style.width = `${Math.abs(x2 - x1)}px`;
            element.style.height = "1px";
          } else {
            element.style.width = "1px";
            element.style.height = `${Math.abs(y2 - y1)}px`;
          }
          formattedResult.appendChild(element);
        }
      }

      if (tableGrid && tableGrid.rows && tableGrid.columns && !layoutRows.length) {
        const tableTop = Number(tableGrid.y || 0) * scale;
        formattedResult.appendChild(createTableElement(tableTop));
        finishFormattedHeight(tableTop + Number(tableGrid.h || 0) * scale);
        return;
      } else if (!tableGrid) {
        let fallbackBottom = 0;
        for (const cell of tableCells) {
          const text = String(cell.text || "").trim();
          if (!text) {
            continue;
          }

          const x = Number(cell.x || 0) * scale;
          const y = Number(cell.y || 0) * scale;
          const width = Math.max(Number(cell.w || 1) * scale, 24);
          const height = Math.max(Number(cell.h || 1) * scale, 12);
          const cellLines = (cell.lines && cell.lines.length ? cell.lines : [cell.text || ""])
            .map((line) => String(line || "").trim())
            .filter(Boolean);

          if (cellLines.length) {
            const longestLine = cellLines.reduce((longest, line) => Math.max(longest, line.length), 1);
            const fontByHeight = (height * 0.82) / Math.max(1, cellLines.length * 1.25);
            const fontByWidth = width / Math.max(4, longestLine * 0.62);
            const fontSize = Math.max(4.5, Math.min(11, fontByHeight, fontByWidth));
            const lineHeight = Math.max(fontSize * 1.25, 6);
            const contentHeight = lineHeight * cellLines.length;
            const element = document.createElement("div");
            applyTextStyle(element, cellLines.join("\\n"), fontSize, {table: true, style: cell.style || {}});
            element.textContent = cellLines.join("\\n");
            element.style.left = `${x}px`;
            element.style.top = `${y + Math.max(1, (height - contentHeight) / 2)}px`;
            element.style.width = `${width}px`;
            element.style.minHeight = `${contentHeight}px`;
            element.style.lineHeight = `${lineHeight}px`;
            element.style.whiteSpace = "pre-wrap";
            element.style.overflow = "visible";
            formattedResult.appendChild(element);
            fallbackBottom = Math.max(fallbackBottom, y + Math.max(height, contentHeight));
          }
        }
        finishFormattedHeight(fallbackBottom);
        if (tableCells.length) {
          return;
        }
      }

      if (layoutRows.length) {
        let lastBottom = 0;
        let layoutYOffset = 0;
        const layoutEntries = layoutRows.map((row) => ({type: "row", y: Number(row.y || 0), row}));
        if (tableGrid && tableGrid.rows && tableGrid.columns) {
          layoutEntries.push({type: "table", y: Number(tableGrid.y || 0)});
        }
        layoutEntries.sort((a, b) => a.y - b.y);

        for (const entry of layoutEntries) {
          if (entry.type === "table") {
            const naturalTop = Number(tableGrid.y || 0) * scale;
            const tableHeight = Number(tableGrid.h || 0) * scale;
            const tableGap = Math.max(6, tableHeight * 0.02);
            const tableTop = Math.max(naturalTop, lastBottom + tableGap);
            formattedResult.appendChild(createTableElement(tableTop));
            lastBottom = tableTop + tableHeight;
            layoutYOffset = tableTop - naturalTop;
            continue;
          }

          const row = entry.row;
          const text = String(row.text || "").trim();
          if (!text) {
            continue;
          }

          const rowHeight = Math.max(Number(row.h || 0), 12);
          const fontSize = Math.max(8, Math.min(22, rowHeight * scale * 0.72));
          const naturalTop = Number(row.y || 0) * scale + layoutYOffset;
          const lineHeight = fontSize * 1.35;
          const minGap = Math.max(3, fontSize * 0.35);
          const top = Math.max(naturalTop, lastBottom + minGap);

          const isCenteredRow = ["heading"].includes(classifyText(text)) || (row.style || {}).align === "center";
          if (isCenteredRow) {
            const element = document.createElement("div");
            applyTextStyle(element, text, fontSize, {style: row.style || {}});
            element.textContent = text;
            element.style.left = "0px";
            element.style.top = `${top}px`;
            element.style.width = `${imageWidth * scale}px`;
            element.style.minHeight = `${lineHeight}px`;
            element.style.overflow = "visible";
            formattedResult.appendChild(element);
            lastBottom = top + Math.max(element.scrollHeight, element.offsetHeight, lineHeight);
            continue;
          }

          let rowBottom = top + lineHeight;
          const segments = row.segments && row.segments.length ? row.segments : [{text, x: row.x, y: row.y, w: row.w, h: row.h, style: row.style}];
          const sortedSegments = [...segments].sort((a, b) => Number(a.x || 0) - Number(b.x || 0));
          if (sortedSegments.length === 1) {
            const segment = sortedSegments[0];
            const segmentText = String(segment.text || "").trim();
            if (!segmentText) {
              continue;
            }
            const segmentX = Number(segment.x || row.x || 0);
            const segmentFontSize = Math.max(8, Math.min(22, Math.max(Number(segment.h || rowHeight), rowHeight) * scale * 0.72));
            const segmentLineHeight = segmentFontSize * 1.32;
            const element = document.createElement("div");
            applyTextStyle(element, segmentText, segmentFontSize, {style: segment.style || row.style || {}});
            element.textContent = segmentText;
            element.style.left = `${Math.max(0, segmentX * scale)}px`;
            element.style.top = `${top}px`;
            element.style.width = `${Math.max(40, (imageWidth - segmentX - 8) * scale)}px`;
            element.style.minHeight = `${segmentLineHeight}px`;
            element.style.whiteSpace = "normal";
            element.style.wordBreak = "normal";
            element.style.overflowWrap = "normal";
            element.style.overflow = "visible";
            formattedResult.appendChild(element);
            rowBottom = top + Math.max(element.scrollHeight, element.offsetHeight, segmentLineHeight);
            lastBottom = rowBottom + Math.max(2, segmentLineHeight * 0.2);
            continue;
          }
          const rowElement = document.createElement("div");
          rowElement.className = "formatted-text";
          rowElement.style.left = "0px";
          rowElement.style.top = `${top}px`;
          rowElement.style.width = `${imageWidth * scale}px`;
          rowElement.style.minHeight = `${lineHeight}px`;
          rowElement.style.whiteSpace = "normal";
          rowElement.style.overflow = "visible";
          rowElement.style.lineHeight = `${lineHeight}px`;
          formattedResult.appendChild(rowElement);
          let maxSegmentHeight = lineHeight;
          const segmentElements = [];
          let occupiedRight = 0;

          for (let index = 0; index < sortedSegments.length; index += 1) {
            const segment = sortedSegments[index];
            const segmentText = String(segment.text || "").trim();
            if (!segmentText) {
              continue;
            }
            const segmentX = Number(segment.x || row.x || 0);
            const segmentFontSize = Math.max(8, Math.min(22, Math.max(Number(segment.h || rowHeight), rowHeight) * scale * 0.72));
            const segmentLineHeight = segmentFontSize * 1.28;
            const nextSegment = sortedSegments[index + 1];
            const segmentWidth = nextSegment
              ? Math.max(24, (Number(nextSegment.x || imageWidth) - segmentX - 3) * scale)
              : Math.max(40, (imageWidth - segmentX - 8) * scale);
            const element = document.createElement("div");
            applyTextStyle(element, segmentText, segmentFontSize, {style: segment.style || row.style || {}});
            element.textContent = segmentText;
            element.style.position = "absolute";
            const desiredLeft = Math.max(0, segmentX * scale);
            element.style.left = `${desiredLeft}px`;
            element.style.top = "0px";
            element.style.width = `${segmentWidth}px`;
            element.style.minHeight = `${segmentLineHeight}px`;
            element.style.whiteSpace = "nowrap";
            element.style.wordBreak = "normal";
            element.style.overflowWrap = "normal";
            element.style.overflow = "visible";
            rowElement.appendChild(element);
            const actualWidth = Math.max(element.scrollWidth, element.offsetWidth, Number(segment.w || 0) * scale);
            const minLeft = occupiedRight + Math.max(3, segmentFontSize * 0.35);
            const finalLeft = desiredLeft < minLeft ? minLeft : desiredLeft;
            if (finalLeft !== desiredLeft) {
              element.style.left = `${finalLeft}px`;
            }
            occupiedRight = finalLeft + actualWidth;
            segmentElements.push({element, fontSize: segmentFontSize, lineHeight: segmentLineHeight});
            maxSegmentHeight = Math.max(maxSegmentHeight, segmentLineHeight);
          }

          const pageWidth = imageWidth * scale;
          const rowOverflow = Math.max(...segmentElements.map((segmentElement) => {
            const left = Number.parseFloat(segmentElement.element.style.left || "0");
            return left + Math.max(segmentElement.element.scrollWidth, segmentElement.element.offsetWidth);
          }), 0);
          if (rowOverflow > pageWidth) {
            const shrinkRatio = Math.max(0.65, pageWidth / rowOverflow);
            maxSegmentHeight = lineHeight;
            for (const segmentElement of segmentElements) {
              const shrunkFontSize = segmentElement.fontSize * shrinkRatio;
              const shrunkLineHeight = segmentElement.lineHeight * shrinkRatio;
              segmentElement.element.style.fontSize = `${shrunkFontSize}px`;
              segmentElement.element.style.lineHeight = `${shrunkLineHeight}px`;
              segmentElement.element.style.minHeight = `${shrunkLineHeight}px`;
              maxSegmentHeight = Math.max(maxSegmentHeight, shrunkLineHeight);
            }
          }

          rowBottom = top + Math.max(rowElement.scrollHeight, rowElement.offsetHeight, maxSegmentHeight);
          lastBottom = rowBottom + Math.max(2, maxSegmentHeight * 0.2);
        }
        finishFormattedHeight(lastBottom);
        return;
      }

      let itemBottom = 0;
      for (const item of items) {
        const rect = normalizeBox(item.box);
        const text = String(item.text || "").trim();
        if (!rect || !text) {
          continue;
        }

        const element = document.createElement("div");
        applyTextStyle(element, text, Math.max(8, rect.height * scale * 0.72), {style: item.style || {}});
        element.textContent = text;
        element.style.left = `${rect.x * scale}px`;
        element.style.top = `${rect.y * scale}px`;
        element.style.width = `${Math.max(rect.width * scale, 40)}px`;
        element.style.minHeight = `${Math.max(rect.height * scale, 10)}px`;
        formattedResult.appendChild(element);
        itemBottom = Math.max(itemBottom, rect.y * scale + Math.max(element.scrollHeight, element.offsetHeight, rect.height * scale));
      }
      finishFormattedHeight(itemBottom);
    }

    async function applySelectedImage(file, dataUrl, statusMessage) {
      selectedFile = file;
      selectedDataUrl = dataUrl;
      clearOverlay();
      clearFormattedResult();
      preview.src = dataUrl;
      preview.style.display = "block";
      placeholder.style.display = "none";
      ocrButton.disabled = false;
      updateExportButtons();
      updatePageNav();
      setStatus(statusMessage || `Đã chọn: ${file.name}`);
    }

    async function prepareSelectedFile(file) {
      const dataUrl = await readFileAsDataUrl(file);
      const isPdf = /\\.pdf$/i.test(file.name) || file.type === "application/pdf";
      if (!isPdf) {
        return [{ filename: file.name, image: dataUrl, autoOrient: true }];
      }

      setStatus(`Đang chuyển PDF: ${file.name}...`);
      const response = await fetch("/prepare-source", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
          filename: file.name,
          file: dataUrl,
          source: "file"
        })
      });
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || "Không thể đọc file PDF.");
      }
      if (Array.isArray(data.pages) && data.pages.length) {
        return data.pages.map((page) => ({
          filename: page.filename || file.name.replace(/\\.pdf$/i, ".png"),
          image: page.image,
          // PDF pages from PyMuPDF are already upright — don't auto-flip.
          autoOrient: false
        }));
      }
      return [{
        filename: data.filename || file.name.replace(/\\.pdf$/i, ".png"),
        image: data.image,
        autoOrient: false
      }];
    }

    async function connectDevice(kind) {
      const label = kind === "printer" ? "máy in" : "máy scan";
      setStatus(`Đang kết nối ${label}...`);
      const response = await fetch(`/devices?kind=${encodeURIComponent(kind)}`);
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || `Không kết nối được ${label}.`);
      }

      const devices = data.devices || [];
      if (!devices.length) {
        setStatus(`Không tìm thấy ${label}.`);
        alert(`Không tìm thấy ${label} trên máy này.`);
        return;
      }

      const names = devices.map((device, index) => `${index + 1}. ${device.name}`).join("\\n");
      const choice = prompt(`Chọn ${label} (nhập số):\\n${names}`, "1");
      if (choice === null) {
        setStatus(`Đã hủy kết nối ${label}.`);
        return;
      }

      const index = Number.parseInt(choice, 10) - 1;
      if (!Number.isInteger(index) || index < 0 || index >= devices.length) {
        setStatus(`Lựa chọn ${label} không hợp lệ.`);
        alert("Số chọn không hợp lệ.");
        return;
      }

      const device = devices[index];
      if (kind === "printer") {
        setStatus(`Đã kết nối máy in: ${device.name}`);
        alert(`Đã kết nối máy in: ${device.name}`);
        return;
      }

      setStatus(`Đang quét ảnh từ máy scan: ${device.name}...`);
      const captureResponse = await fetch("/capture-device", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({ kind, device_id: device.id, device_name: device.name })
      });
      const captureData = await captureResponse.json();
      if (!captureResponse.ok || !captureData.ok) {
        throw new Error(captureData.error || "Không lấy được ảnh từ máy scan.");
      }

      const capturedName = captureData.filename || "scanner_capture.png";
      const capturedFile = new File([""], capturedName, { type: "image/png" });
      await applySelectedImage(
        capturedFile,
        captureData.image,
        `Đã quét từ máy scan: ${device.name}`
      );
      selectedFiles = [{
        file: capturedFile,
        image: captureData.image,
        displayImage: captureData.image,
        name: capturedName,
        ocrData: null
      }];
      currentPageIndex = 0;
      updatePageNav();
    }

    sourceSelect.addEventListener("change", async () => {
      const value = sourceSelect.value;
      sourceSelect.value = "";
      if (!value) {
        return;
      }

      try {
        if (value === "file") {
          fileInput.value = "";
          fileInput.click();
          return;
        }
        if (value === "printer" || value === "scanner") {
          await connectDevice(value);
        }
      } catch (error) {
        setStatus(String(error.message || error));
        alert(String(error.message || error));
      }
    });

    fileInput.addEventListener("change", async () => {
      const files = Array.from(fileInput.files || []);
      clearOverlay();
      clearResultData();
      clearFormattedResult();
      updateExportButtons();

      if (!files.length) {
        selectedFile = null;
        selectedFiles = [];
        currentPageIndex = 0;
        selectedDataUrl = "";
        preview.style.display = "none";
        placeholder.style.display = "block";
        ocrButton.disabled = true;
        updatePageNav();
        setStatus("Chọn ảnh/pdf để bắt đầu.");
        return;
      }

      try {
        const preparedList = [];
        for (const file of files) {
          const preparedPages = await prepareSelectedFile(file);
          for (const prepared of preparedPages) {
            preparedList.push({
              file: new File([file], prepared.filename, { type: "image/png" }),
              image: prepared.image,
              displayImage: prepared.image,
              name: prepared.filename,
              autoOrient: prepared.autoOrient !== false,
              ocrData: null
            });
          }
        }
        if (!preparedList.length) {
          throw new Error("Không đọc được trang ảnh nào từ file đã chọn.");
        }
        if (isMau2MultiPage() && preparedList.length < 2) {
          throw new Error("Mẫu số 2 cần ít nhất 2 trang (2 ảnh, hoặc 1 PDF có ≥ 2 trang).");
        }
        selectedFiles = preparedList;
        currentPageIndex = 0;
        selectedFile = preparedList[0].file;
        const mau2Hint = isMau2MultiPage()
          ? `Mẫu số 2: đã chọn ${preparedList.length} trang (PDF/ảnh). OCR rồi Export JSON sẽ gộp các trang.`
          : undefined;
        await applySelectedImage(
          preparedList[0].file,
          preparedList[0].image,
          preparedList.length > 1
            ? (mau2Hint || `Đã chọn ${preparedList.length} trang. Dùng ‹ › để xem từng trang, rồi OCR.`)
            : undefined
        );
        updatePageNav();
      } catch (error) {
        selectedFile = null;
        selectedFiles = [];
        currentPageIndex = 0;
        selectedDataUrl = "";
        preview.style.display = "none";
        placeholder.style.display = "block";
        ocrButton.disabled = true;
        updatePageNav();
        updateExportButtons();
        setStatus(String(error.message || error));
        alert(String(error.message || error));
      }
    });

    ocrButton.addEventListener("click", async () => {
      const filesToOcr = selectedFiles.length
        ? selectedFiles
        : (selectedFile && selectedDataUrl ? [{ file: selectedFile, image: selectedDataUrl, displayImage: selectedDataUrl, name: selectedFile.name, ocrData: null }] : []);
      if (!filesToOcr.length) {
        alert("Vui lòng chọn ảnh/pdf hoặc kết nối máy scan trước.");
        return;
      }
      if (isMau2MultiPage() && filesToOcr.length < 2) {
        alert("Mẫu số 2 cần ít nhất 2 trang (2 ảnh, hoặc 1 PDF có ≥ 2 trang).");
        return;
      }

      ocrButton.disabled = true;
      exportDocx.disabled = true;
      exportXlsx.disabled = true;
      exportJson.disabled = true;
      prevPage.disabled = true;
      nextPage.disabled = true;
      clearOverlay();
      clearResultData();
      clearFormattedResult();
      setStatus(engine.value === "vietocr" ? "Đang OCR bằng VietOCR..." : "Đang OCR bằng PaddleOCR...");

      try {
        for (let index = 0; index < filesToOcr.length; index += 1) {
          const item = filesToOcr[index];
          currentPageIndex = index;
          selectedFile = item.file;
          selectedDataUrl = item.displayImage || item.image;
          preview.src = selectedDataUrl;
          preview.style.display = "block";
          placeholder.style.display = "none";
          updatePageNav();
          setStatus(`Đang OCR ảnh ${index + 1}/${filesToOcr.length}: ${item.name}...`);

          const response = await fetch("/ocr", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({
              filename: item.name,
              image: item.image,
              engine: engine.value,
              auto_orient: item.autoOrient !== false
            })
          });
          const data = await response.json();
          if (!response.ok || !data.ok) {
            throw new Error(data.error || `OCR lỗi ở ảnh ${item.name}.`);
          }

          if (data.oriented_image && data.orientation_angle) {
            item.displayImage = data.oriented_image;
            selectedDataUrl = data.oriented_image;
            preview.src = data.oriented_image;
          } else {
            item.displayImage = item.image;
          }
          // Persist OCR before any UI wait so later pages are never skipped.
          item.ocrData = data;
          selectedFiles[index] = item;
          updatePageNav();
        }

        await showPage(0);
        const totalSeconds = selectedFiles.reduce((sum, page) => sum + Number(page.ocrData?.elapsed_seconds || 0), 0);
        const doneCount = selectedFiles.filter((page) => page.ocrData).length;
        setStatus(`OCR hoàn tất ${doneCount}/${selectedFiles.length} ảnh trong ${totalSeconds.toFixed(2)}s. Dùng ‹ › để xem từng trang.`);
      } catch (error) {
        latestText = String(error.message || error);
        setStatus("OCR lỗi.");
      } finally {
        ocrButton.disabled = !selectedFiles.length;
        updatePageNav();
        syncExportFromPages();
      }
    });

    async function exportResult(format) {
      const exportPages = collectExportPages();
      if (!exportPages.length) {
        alert("Chưa có kết quả OCR để export.");
        return;
      }

      const page = currentPage();
      const baseName = page
        ? page.name.replace(/\\.[^.]+$/, "")
        : (selectedFile ? selectedFile.name.replace(/\\.[^.]+$/, "") : "ocr_result");
      const multiSuffix = exportPages.length > 1 ? `_all${exportPages.length}` : "";
      const exportName = isMau2MultiPage() && exportPages.length > 1
        ? `${baseName}_mau2_ocr.${format}`
        : `${baseName}${multiSuffix}_ocr.${format}`;

      let text = exportPages.map((item) => item.text).join("\\n\\n");
      let tableRows = exportPages.flatMap((item) => item.table_rows || []);
      // Mẫu 2 JSON: luôn gộp toàn bộ trang đã OCR (kể cả PDF nhiều trang).
      if (format === "json" && isMau2MultiPage()) {
        const mau2Pages = mau2PagesForJson();
        if (mau2Pages.length < 2) {
          alert("Mẫu số 2 cần OCR ít nhất 2 trang trước khi Export JSON.");
          return;
        }
        text = mau2Pages.map((item) => item.ocrData.text || "").filter(Boolean).join("\\n\\n");
        tableRows = mau2Pages.flatMap((item) => item.ocrData.table_rows || []);
      }

      const payload = {
        text,
        table_rows: tableRows,
        formatted_data: exportPages[0].formatted_data || latestFormattedData,
        json_type: jsonType.value,
        filename: exportName
      };
      if (format === "docx" || format === "xlsx") {
        payload.pages = exportPages;
      }

      const response = await fetch(`/export?format=${encodeURIComponent(format)}`, {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        alert(data.error || "Export lỗi.");
        return;
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = exportName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }

    exportDocx.addEventListener("click", () => exportResult("docx"));
    exportXlsx.addEventListener("click", () => exportResult("xlsx"));
    exportJson.addEventListener("click", () => exportResult("json"));
    preview.addEventListener("load", () => drawOverlay(latestBoxes, latestTableLines));
    window.addEventListener("resize", () => {
      drawOverlay(latestBoxes, latestTableLines);
      if (latestFormattedData) {
        renderFormattedResult(latestFormattedData);
      }
    });

    async function pollStatus() {
      try {
        const response = await fetch("/status");
        const data = await response.json();
        if (data.loading) {
          setStatus(`Đang nạp model ${data.engine}...`);
        } else if (data.error) {
          setStatus(`Lỗi nạp model: ${data.error}`);
        } else if (data.engine) {
          setStatus(`Model ${data.engine} đã sẵn sàng.`);
        }
      } catch (_) {
        // Ignore transient status polling errors.
      }
    }

    setInterval(pollStatus, 2500);
    pollStatus();
  </script>
</body>
</html>
"""


class OCRWebApp:
    def __init__(self):
        self.recognitor = None
        self.detector = None
        self.current_engine = None
        self.loading = False
        self.load_error = None
        self.model_lock = threading.Lock()

    def preload(self, engine="vietocr"):
        def _run():
            try:
                self.ensure_models(engine)
            except Exception:
                # ensure_models already stores load_error for /status
                pass

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()

    def ensure_models(self, engine):
        with self.model_lock:
            model_ready = (
                self.current_engine == engine
                and self.detector is not None
                and (engine == "paddle" or self.recognitor is not None)
            )
            if model_ready:
                return

            self.loading = True
            self.load_error = None
            try:
                if engine == "paddle":
                    from predict import create_fast_model

                    self.recognitor = None
                    self.detector = create_fast_model(
                        quiet=True,
                        det_limit_side_len=OCR_DET_LIMIT_SIDE,
                    )
                else:
                    from predict import create_models

                    self.recognitor, self.detector = create_models(
                        quiet=True,
                        beamsearch=False,
                        det_limit_side_len=OCR_DET_LIMIT_SIDE,
                    )

                self.current_engine = engine
            except Exception as exc:
                self.load_error = str(exc)
                self.current_engine = None
                self.recognitor = None
                self.detector = None
                raise
            finally:
                self.loading = False

    def ocr(self, image_path, engine):
        self.ensure_models(engine)
        if engine == "paddle":
            from predict import predict_fast

            boxes, texts = predict_fast(
                self.detector,
                str(image_path),
                None,
                verbose=False,
                draw_output=False,
            )
        else:
            from predict import predict

            boxes, texts = predict(
                self.recognitor,
                self.detector,
                str(image_path),
                None,
                padding=6,
                verbose=False,
                draw_output=False,
            )

        return boxes, texts


APP_STATE = OCRWebApp()


def serialize_box(box):
    if not isinstance(box, (list, tuple)):
        return []
    serialized = []
    for point in box:
        if isinstance(point, (list, tuple)) and len(point) >= 2:
            serialized.append([float(point[0]), float(point[1])])
    return serialized


def box_bounds(box):
    points = serialize_box(box)
    if len(points) == 2:
        x1, y1 = points[0]
        x2, y2 = points[1]
        points = [[x1, y1], [x2, y1], [x2, y2], [x1, y2]]
    if not points:
        return None

    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return {
        "x": min(xs),
        "y": min(ys),
        "w": max(xs) - min(xs),
        "h": max(ys) - min(ys),
        "cx": (min(xs) + max(xs)) / 2,
        "cy": (min(ys) + max(ys)) / 2,
    }


def split_choice_cells(text):
    parts = [part.strip() for part in re.split(r"(?=\b[A-D]\s*[\.\)])", text) if part.strip()]
    return parts if len(parts) >= 2 else None


def vertical_overlap_ratio(bounds_a, bounds_b):
    top = max(bounds_a["y"], bounds_b["y"])
    bottom = min(bounds_a["y"] + bounds_a["h"], bounds_b["y"] + bounds_b["h"])
    overlap = max(0, bottom - top)
    min_height = max(1, min(bounds_a["h"], bounds_b["h"]))
    return overlap / min_height


def group_items_by_text_rows(items):
    rows = []
    for item in sorted(items, key=lambda entry: (entry["bounds"]["y"], entry["bounds"]["x"])):
        target_row = None
        for row in rows:
            row_bounds = row["bounds"]
            center_delta = abs(row["cy"] - item["bounds"]["cy"])
            row_heights = sorted(row_item["bounds"]["h"] for row_item in row["items"] if row_item["bounds"]["h"] > 0)
            row_height = row_heights[len(row_heights) // 2] if row_heights else row_bounds["h"]
            comparable_height = max(1, min(row_height, item["bounds"]["h"]))
            max_center_delta = max(4, comparable_height * 0.32)
            overlaps = vertical_overlap_ratio(row_bounds, item["bounds"]) >= 0.55
            if overlaps or center_delta <= max_center_delta:
                target_row = row
                break

        if target_row is None:
            rows.append(
                {
                    "cy": item["bounds"]["cy"],
                    "bounds": dict(item["bounds"]),
                    "items": [item],
                }
            )
        else:
            target_row["items"].append(item)
            all_bounds = [row_item["bounds"] for row_item in target_row["items"]]
            x1 = min(bounds["x"] for bounds in all_bounds)
            y1 = min(bounds["y"] for bounds in all_bounds)
            x2 = max(bounds["x"] + bounds["w"] for bounds in all_bounds)
            y2 = max(bounds["y"] + bounds["h"] for bounds in all_bounds)
            target_row["bounds"] = {
                "x": x1,
                "y": y1,
                "w": x2 - x1,
                "h": y2 - y1,
                "cx": (x1 + x2) / 2,
                "cy": (y1 + y2) / 2,
            }
            target_row["cy"] = target_row["bounds"]["cy"]

    return sorted(rows, key=lambda row: row["bounds"]["y"])


def build_table_rows(boxes, texts):
    items = []
    for box, text in zip(boxes, texts):
        text = str(text).strip()
        bounds = box_bounds(box)
        if text and bounds:
            items.append({"text": text, "bounds": bounds})

    if not items:
        return [[line] for line in texts if str(line).strip()]

    table_rows = []
    for row in group_items_by_text_rows(items):
        row_items = sorted(row["items"], key=lambda entry: entry["bounds"]["x"])
        row_text = " ".join(item["text"] for item in row_items).strip()
        choice_cells = split_choice_cells(row_text)
        if choice_cells:
            table_rows.append(choice_cells)
        elif len(row_items) > 1:
            table_rows.append([item["text"] for item in row_items])
        else:
            table_rows.append([row_text])

    return table_rows


def estimate_text_styles(image_path, boxes, texts):
    import cv2
    import numpy as np
    from predict import imread_unicode

    image = imread_unicode(str(image_path))
    heights = []
    bounds_by_index = {}
    for index, box in enumerate(boxes):
        bounds = box_bounds(box)
        if bounds:
            bounds_by_index[index] = bounds
            if bounds["h"] > 0:
                heights.append(bounds["h"])

    median_height = sorted(heights)[len(heights) // 2] if heights else 16
    image_height, image_width = image.shape[:2] if image is not None else (1, 1)
    styles = {}

    for index, text in enumerate(texts):
        text = str(text).strip()
        bounds = bounds_by_index.get(index)
        if not text or not bounds:
            styles[index] = {"color": "#111827", "bold": False, "align": "left"}
            continue

        color = "#111827"
        if image is not None:
            x1 = max(0, int(bounds["x"]) - 2)
            y1 = max(0, int(bounds["y"]) - 2)
            x2 = min(image_width, int(bounds["x"] + bounds["w"]) + 2)
            y2 = min(image_height, int(bounds["y"] + bounds["h"]) + 2)
            crop = image[y1:y2, x1:x2]
            if crop.size:
                rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
                hsv = cv2.cvtColor(crop, cv2.COLOR_BGR2HSV)
                saturation = hsv[:, :, 1]
                value = hsv[:, :, 2]
                red = rgb[:, :, 0]
                green = rgb[:, :, 1]
                blue = rgb[:, :, 2]
                red_mask = (red > 120) & (red > green * 1.25) & (red > blue * 1.25) & (saturation > 45)
                dark_mask = value < 115
                foreground = red_mask | dark_mask
                if np.any(red_mask) and np.count_nonzero(red_mask) >= max(3, np.count_nonzero(foreground) * 0.25):
                    avg = rgb[red_mask].mean(axis=0)
                    color = "#{:02x}{:02x}{:02x}".format(
                        int(min(220, max(150, avg[0]))),
                        int(min(80, max(20, avg[1]))),
                        int(min(80, max(20, avg[2]))),
                    )

        letters = [char for char in text if char.isalpha()]
        upper_ratio = sum(1 for char in letters if char.upper() == char) / len(letters) if letters else 0
        is_large = bounds["h"] >= median_height * 1.25
        is_heading_like = upper_ratio >= 0.65 and len(text) >= 8
        is_centered = abs(bounds["cx"] - image_width / 2) <= image_width * 0.18

        styles[index] = {
            "color": color,
            "bold": bool(is_large or is_heading_like),
            "align": "center" if is_centered and (is_large or is_heading_like or bounds["y"] < image_height * 0.25) else "left",
        }

    return styles


def merge_styles(styles):
    if not styles:
        return {"color": "#111827", "bold": False, "align": "left"}

    red_styles = [style for style in styles if style.get("color", "#111827").lower() != "#111827"]
    center_count = sum(1 for style in styles if style.get("align") == "center")
    bold_count = sum(1 for style in styles if style.get("bold"))
    return {
        "color": red_styles[0]["color"] if red_styles else "#111827",
        "bold": bold_count >= max(1, len(styles) // 2),
        "align": "center" if center_count >= max(1, len(styles) // 2) else "left",
    }


def build_layout_rows(boxes, texts, skip_indices=None, styles=None):
    skip_indices = skip_indices or set()
    styles = styles or {}
    items = []
    for index, (box, text) in enumerate(zip(boxes, texts)):
        if index in skip_indices:
            continue
        text = str(text).strip()
        bounds = box_bounds(box)
        if text and bounds:
            items.append({"index": index, "text": text, "bounds": bounds, "style": styles.get(index, {})})

    if not items:
        return []

    heights = sorted(item["bounds"]["h"] for item in items if item["bounds"]["h"] > 0)
    median_height = heights[len(heights) // 2] if heights else 16

    layout_rows = []
    for row in group_items_by_text_rows(items):
        row_items = sorted(row["items"], key=lambda entry: entry["bounds"]["x"])
        x1 = min(item["bounds"]["x"] for item in row_items)
        y1 = min(item["bounds"]["y"] for item in row_items)
        x2 = max(item["bounds"]["x"] + item["bounds"]["w"] for item in row_items)
        y2 = max(item["bounds"]["y"] + item["bounds"]["h"] for item in row_items)
        text = " ".join(item["text"] for item in row_items).strip()
        if text:
            row_style = merge_styles([item.get("style", {}) for item in row_items])
            segments = [
                {
                    "text": item["text"],
                    "x": item["bounds"]["x"],
                    "y": item["bounds"]["y"],
                    "w": item["bounds"]["w"],
                    "h": item["bounds"]["h"],
                    "style": item.get("style", {}),
                }
                for item in row_items
            ]
            layout_rows.append(
                {
                    "text": text,
                    "x": x1,
                    "y": y1,
                    "w": x2 - x1,
                    "h": max(y2 - y1, median_height),
                    "style": row_style,
                    "segments": segments,
                }
            )

    return layout_rows


def table_axes(table_lines):
    horizontal = []
    vertical = []
    for line in table_lines:
        if len(line) < 4:
            continue
        x1, y1, x2, y2 = [float(value) for value in line]
        if abs(y2 - y1) <= abs(x2 - x1):
            horizontal.append((min(x1, x2), (y1 + y2) / 2, max(x1, x2)))
        else:
            vertical.append(((x1 + x2) / 2, min(y1, y2), max(y1, y2)))

    xs = sorted({round(x, 1) for x, _y1, _y2 in vertical})
    ys = sorted({round(y, 1) for _x1, y, _x2 in horizontal})
    return xs, ys


def build_table_cells(boxes, texts, table_lines, styles=None):
    styles = styles or {}
    xs, ys = table_axes(table_lines)
    if len(xs) < 2 or len(ys) < 2:
        return [], set(), None

    table_left, table_right = min(xs), max(xs)
    table_top, table_bottom = min(ys), max(ys)
    table_grid = {
        "x": table_left,
        "y": table_top,
        "w": table_right - table_left,
        "h": table_bottom - table_top,
        "columns": [xs[index + 1] - xs[index] for index in range(len(xs) - 1)],
        "rows": [ys[index + 1] - ys[index] for index in range(len(ys) - 1)],
    }
    cell_items = {}
    assigned_indices = set()

    def best_axis_interval(start, end, axes):
        best_index = None
        best_overlap = 0
        for idx in range(len(axes) - 1):
            overlap = max(0, min(end, axes[idx + 1]) - max(start, axes[idx]))
            if overlap > best_overlap:
                best_overlap = overlap
                best_index = idx
        return best_index, best_overlap

    for index, (box, text) in enumerate(zip(boxes, texts)):
        text = str(text).strip()
        bounds = box_bounds(box)
        if not text or not bounds:
            continue

        x1 = bounds["x"]
        y1 = bounds["y"]
        x2 = bounds["x"] + bounds["w"]
        y2 = bounds["y"] + bounds["h"]
        if x2 < table_left or x1 > table_right or y2 < table_top or y1 > table_bottom:
            continue

        col, col_overlap = best_axis_interval(x1, x2, xs)
        row, row_overlap = best_axis_interval(y1, y2, ys)
        if col is None or row is None or col_overlap <= 0 or row_overlap <= 0:
            continue

        cell_items.setdefault((row, col), []).append({"index": index, "text": text, "bounds": bounds, "style": styles.get(index, {})})
        assigned_indices.add(index)

    table_cells = []
    for (row, col), items in sorted(cell_items.items()):
        items = sorted(items, key=lambda item: (item["bounds"]["cy"], item["bounds"]["x"]))
        cell_style = merge_styles([item.get("style", {}) for item in items])
        heights = sorted(item["bounds"]["h"] for item in items if item["bounds"]["h"] > 0)
        median_height = heights[len(heights) // 2] if heights else 10
        line_tolerance = max(5, median_height * 0.65)
        item_lines = []
        for item in items:
            target_line = None
            for line_index, line_items in enumerate(item_lines):
                line_cy = sum(line_item["bounds"]["cy"] for line_item in line_items) / len(line_items)
                if abs(line_cy - item["bounds"]["cy"]) <= line_tolerance:
                    target_line = line_index
                    break
            if target_line is None:
                item_lines.append([item])
            else:
                item_lines[target_line].append(item)

        indexed_items = []
        cell_lines = []
        for line_index, line_items in enumerate(item_lines):
            sorted_line_items = sorted(line_items, key=lambda line_item: line_item["bounds"]["x"])
            line_text = " ".join(item["text"] for item in sorted_line_items).strip()
            if line_text:
                cell_lines.append(line_text)
            for item in sorted_line_items:
                indexed_items.append((line_index, item))

        table_cells.append(
            {
                "row": row,
                "col": col,
                "text": " ".join(item["text"] for item in items),
                "x": xs[col],
                "y": ys[row],
                "w": max(1, xs[col + 1] - xs[col]),
                "h": max(1, ys[row + 1] - ys[row]),
                "style": cell_style,
                "lines": cell_lines,
                "items": [
                    {
                        "text": item["text"],
                        "line_index": line_index,
                        "x": max(0, item["bounds"]["x"] - xs[col]),
                        "w": item["bounds"]["w"],
                        "h": item["bounds"]["h"],
                        "style": item.get("style", {}),
                    }
                    for line_index, item in indexed_items
                ],
            }
        )

    return table_cells, assigned_indices, table_grid


def merge_line_segments(lines, orientation, tolerance=8):
    if not lines:
        return []

    if orientation == "horizontal":
        lines = sorted(lines, key=lambda line: (line[1], line[0]))
    else:
        lines = sorted(lines, key=lambda line: (line[0], line[1]))

    merged = []
    for line in lines:
        x1, y1, x2, y2 = line
        placed = False
        for idx, current in enumerate(merged):
            cx1, cy1, cx2, cy2 = current
            if orientation == "horizontal":
                same_axis = abs(cy1 - y1) <= tolerance
                overlaps = x1 <= cx2 + tolerance and x2 >= cx1 - tolerance
                if same_axis and overlaps:
                    new_y = int(round((cy1 + y1) / 2))
                    merged[idx] = [min(cx1, x1), new_y, max(cx2, x2), new_y]
                    placed = True
                    break
            else:
                same_axis = abs(cx1 - x1) <= tolerance
                overlaps = y1 <= cy2 + tolerance and y2 >= cy1 - tolerance
                if same_axis and overlaps:
                    new_x = int(round((cx1 + x1) / 2))
                    merged[idx] = [new_x, min(cy1, y1), new_x, max(cy2, y2)]
                    placed = True
                    break

        if not placed:
            merged.append(line)

    return merged


def detect_table_lines(image_path):
    import cv2
    import numpy as np
    from predict import imread_unicode

    image = imread_unicode(str(image_path))
    if image is None:
        return []

    height, width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        25,
        12,
    )

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width // 18), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, height // 28)))
    horizontal_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    vertical_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)

    horizontal_lines = []
    contours, _ = cv2.findContours(horizontal_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w >= width * 0.18 and h <= max(12, height * 0.02):
            horizontal_lines.append([x, y + h // 2, x + w, y + h // 2])

    vertical_lines = []
    contours, _ = cv2.findContours(vertical_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if h >= height * 0.04 and w <= max(12, width * 0.02):
            vertical_lines.append([x + w // 2, y, x + w // 2, y + h])

    horizontal_lines = merge_line_segments(horizontal_lines, "horizontal")
    vertical_lines = merge_line_segments(vertical_lines, "vertical")

    if not horizontal_lines or not vertical_lines:
        return horizontal_lines + vertical_lines

    # Keep lines around real grid intersections; this removes most page borders and text strokes.
    filtered_horizontal = []
    for line in horizontal_lines:
        x1, y1, x2, _y2 = line
        intersections = sum(1 for vx1, vy1, _vx2, vy2 in vertical_lines if x1 - 8 <= vx1 <= x2 + 8 and vy1 - 8 <= y1 <= vy2 + 8)
        if intersections >= 2:
            filtered_horizontal.append(line)

    filtered_vertical = []
    for line in vertical_lines:
        x1, y1, _x2, y2 = line
        intersections = sum(1 for hx1, hy1, hx2, _hy2 in horizontal_lines if y1 - 8 <= hy1 <= y2 + 8 and hx1 - 8 <= x1 <= hx2 + 8)
        if intersections >= 2:
            filtered_vertical.append(line)

    return filtered_horizontal + filtered_vertical


def safe_filename(filename, allowed_extensions=None):
    allowed = allowed_extensions or IMAGE_EXTENSIONS
    name = Path(filename or "upload.png").name
    suffix = Path(name).suffix.lower()
    if suffix not in allowed:
        suffix = ".png"
    stem = Path(name).stem or "upload"
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-") or "upload"
    stem = stem[:80]
    timestamp = int(time.time() * 1000)
    return f"{stem}_{timestamp}{suffix}"


def decode_data_url(data_url):
    payload = data_url or ""
    if "," in payload:
        _header, payload = payload.split(",", 1)
    return base64.b64decode(payload)


def image_bytes_to_data_url(image_bytes, mime="image/png"):
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def convert_pdf_bytes_to_pages(pdf_bytes, zoom=2.0):
    """Rasterize every PDF page to PNG bytes (same scale as single-page helper)."""
    try:
        import fitz
    except ImportError as exc:
        raise ValueError(
            "Chưa hỗ trợ PDF. Cài thêm pymupdf rồi thử lại: pip install pymupdf"
        ) from exc

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if document.page_count < 1:
            raise ValueError("File PDF không có trang nào.")
        matrix = fitz.Matrix(zoom, zoom)
        pages = []
        for index in range(document.page_count):
            page = document.load_page(index)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            pages.append(pixmap.tobytes("png"))
        return pages
    finally:
        document.close()


def convert_pdf_bytes_to_png(pdf_bytes):
    return convert_pdf_bytes_to_pages(pdf_bytes)[0]


def pdf_pages_as_named_pngs(filename, pdf_bytes):
    pages = convert_pdf_bytes_to_pages(pdf_bytes)
    stem = Path(filename or "document").stem or "document"
    named = []
    for index, png_bytes in enumerate(pages, start=1):
        page_name = f"{stem}_p{index}.png" if len(pages) > 1 else f"{stem}.png"
        named.append((page_name, png_bytes))
    return named


def list_windows_printers():
    import subprocess

    command = (
        "Get-Printer | Select-Object Name, DriverName, PortName | "
        "ConvertTo-Json -Compress"
    )
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-Command", command],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout or "Không đọc được danh sách máy in.").strip()
        raise ValueError(message)

    raw = (completed.stdout or "").strip()
    if not raw:
        return []

    data = json.loads(raw)
    if isinstance(data, dict):
        data = [data]

    devices = []
    for index, item in enumerate(data or []):
        name = str(item.get("Name") or "").strip()
        if not name:
            continue
        devices.append(
            {
                "id": name,
                "name": name,
                "driver": str(item.get("DriverName") or ""),
                "port": str(item.get("PortName") or ""),
                "index": index,
            }
        )
    return devices


def list_windows_scanners():
    import subprocess

    script = r"""
$ErrorActionPreference = 'Stop'
try {
  $manager = New-Object -ComObject WIA.DeviceManager
  $devices = @()
  foreach ($info in $manager.DeviceInfos) {
    if ($info.Type -eq 1) {
      $name = $null
      try { $name = $info.Properties('Name').Value } catch {}
      if (-not $name) { $name = $info.DeviceID }
      $devices += [PSCustomObject]@{ id = $info.DeviceID; name = $name }
    }
  }
  $devices | ConvertTo-Json -Compress
} catch {
  Write-Output '[]'
}
"""
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout or "Không đọc được danh sách máy scan.").strip()
        raise ValueError(message)

    raw = (completed.stdout or "").strip()
    if not raw:
        return []

    data = json.loads(raw)
    if isinstance(data, dict):
        data = [data]

    devices = []
    for index, item in enumerate(data or []):
        device_id = str(item.get("id") or "").strip()
        name = str(item.get("name") or device_id).strip()
        if not device_id and not name:
            continue
        devices.append({"id": device_id or name, "name": name or device_id, "index": index})
    return devices


def capture_scanner_image(device_id):
    import subprocess
    import tempfile

    if not device_id:
        raise ValueError("Chưa chọn máy scan.")

    with tempfile.TemporaryDirectory(prefix="ocr_scan_") as temp_dir:
        output_path = Path(temp_dir) / "scan.png"
        script = f"""
$ErrorActionPreference = 'Stop'
$deviceId = {json.dumps(device_id)}
$outputPath = {json.dumps(str(output_path))}
$manager = New-Object -ComObject WIA.DeviceManager
$info = $null
foreach ($item in $manager.DeviceInfos) {{
  if ($item.DeviceID -eq $deviceId) {{ $info = $item; break }}
}}
if (-not $info) {{ throw 'Không tìm thấy máy scan đã chọn.' }}
$device = $info.Connect()
$item = $device.Items(1)
$image = $item.Transfer("{{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}}")
$image.SaveFile($outputPath)
"""
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode != 0 or not output_path.exists():
            message = (completed.stderr or completed.stdout or "Quét ảnh thất bại.").strip()
            raise ValueError(message)
        return output_path.read_bytes()


def get_local_ip():
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        return "127.0.0.1"


def get_image_size(image_path):
    from PIL import Image

    with Image.open(image_path) as image:
        return image.size


def score_detection_orientation(image, detector):
    import cv2

    height, width = image.shape[:2]
    longest = max(height, width)
    scale = min(1.0, 1000 / longest) if longest else 1.0
    sample = (
        cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        if scale < 0.999
        else image
    )
    try:
        result = detector.ocr(sample, cls=False, det=True, rec=False)
    except Exception:
        return 0.0

    lines = result[0] if result and result[0] else []
    if not lines:
        return 0.0

    score = 0.0
    for line in lines:
        points = line
        if (
            isinstance(line, (list, tuple))
            and line
            and isinstance(line[0], (list, tuple))
            and line[0]
            and isinstance(line[0][0], (list, tuple))
        ):
            points = line[0]
        if not (
            isinstance(points, (list, tuple))
            and len(points) >= 4
            and isinstance(points[0], (list, tuple))
        ):
            continue
        xs = [float(point[0]) for point in points[:4]]
        ys = [float(point[1]) for point in points[:4]]
        box_w = max(xs) - min(xs)
        box_h = max(ys) - min(ys)
        aspect = box_w / max(box_h, 1.0)
        score += 1.0
        if aspect >= 1.6:
            score += 2.5
        elif aspect >= 1.2:
            score += 1.0
        score += min(box_w, 240.0) / 120.0
    return score


def count_document_signals(text):
    value = str(text or "")
    patterns = (
        r"CỘNG\s*HÒA\s*XÃ\s*HỘI",
        r"CĂN\s*CƯỚC|Citizen\s*Identity\s*Card",
        r"CỤC\s*TRƯỞNG|DIRECTOR\s*GENERAL",
        r"GIẤY\s*CHỨNG\s*NHẬN|CHỨNG\s*NHẬN\s*QUYỀN",
        r"Th[ửừứùa]\s*đ[ấaá]t\s*s[ốôo]",
        r"T[ờò]\s*b[ảàáa]n\s*đ[ồôo]",
        r"Di[ệe]n\s*t[íiì]ch\s*[:\.]?\s*[0-9]",
        r"M[ụu]c\s*đ[íi]ch\s*s[ửủ]\s*d[ụu]ng|Lo[ạa]i\s*đ[ấaá]t",
        r"Họ\s*v[àa]\s*tên|Full\s*name|H[ộo]\s*[oôóòọ]ng",
        r"(?<!\d)\d{9,12}(?!\d)",
        r"Ngày\s*,?\s*tháng\s*,?\s*năm|Date\s*,?\s*month\s*,?\s*year",
        r"Địa\s*ch[ỉi]\s*th[ửừứùa]|Nơi\s*thường\s*trú|Place\s*of\s*residence",
    )
    return sum(1 for pattern in patterns if re.search(pattern, value, re.I))


def has_strong_upright_cccd_signals(text):
    """True when OCR looks like a readable upright CCCD (not sideways junk)."""
    value = str(text or "")
    has_id = bool(re.search(r"(?<!\d)\d{9,12}(?!\d)", value))
    has_title = bool(
        re.search(r"CĂN\s*CƯỚC|Citizen\s*Identity\s*Card|CHỨNG\s*NHẬN", value, re.I)
    )
    has_name_label = bool(re.search(r"Họ\s*v[àa]\s*tên|Full\s*name", value, re.I))
    return has_id and (has_title or has_name_label)


def score_ocr_text_quality(text):
    value = str(text or "")
    if not value.strip():
        return -100.0

    score = 0.0
    signals = count_document_signals(value)
    score += 14.0 * signals

    if re.search(r"(?<!\d)\d{9,12}(?!\d)", value):
        score += 18.0
    if re.search(r"\d{1,2}/\d{1,2}/\d{4}", value):
        score += 8.0
    if re.search(r"CĂN\s*CƯỚC|Citizen\s*Identity\s*Card|CHỨNG\s*NHẬN|GIẤY\s*CHỨNG", value, re.I):
        score += 12.0
    if re.search(
        r"CỤC\s*TRƯỞNG|DIRECTOR\s*GENERAL|IDVNM|Ngón\s*trỏ|Personal\s*identification|"
        r"Đặc\s*điểm|Date\s*,?\s*month\s*,?\s*year|Ngày\s*,?\s*tháng\s*,?\s*năm",
        value,
        re.I,
    ):
        score += 10.0
    if re.search(
        r"Th[ửừứùa]\s*đ[ấaá]t|T[ờò]\s*b[ảàáa]n\s*đ[ồôo]|M[ụu]c\s*đ[íi]ch\s*s[ửủ]\s*d[ụu]ng|"
        r"Lo[ạa]i\s*đ[ấaá]t|Di[ệe]n\s*t[íiì]ch|H[ộo]\s*[oôóòọ]ng",
        value,
        re.I,
    ):
        score += 18.0
    if re.search(
        r"Họ\s*v[àa]\s*tên|Full\s*name|Giới\s*tính|Sex|Ngày\s*sinh|Date\s*of\s*birth|"
        r"Nơi\s*thường\s*trú|Place\s*of\s*residence|Địa\s*chỉ",
        value,
        re.I,
    ):
        score += 8.0

    vietnamese_chars = re.findall(
        r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ"
        r"ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]",
        value,
    )
    score += min(12.0, float(len(vietnamese_chars)) * 0.35)

    # Upside-down / sideways OCR tends to produce many tiny junk tokens/lines.
    tokens = [token for token in re.split(r"\s+", value) if token]
    short_tokens = sum(1 for token in tokens if len(token) <= 2)
    if tokens:
        score -= 12.0 * (short_tokens / len(tokens))

    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if lines:
        tiny_lines = sum(1 for line in lines if len(line) <= 2)
        score -= 18.0 * (tiny_lines / len(lines))

    # Strong penalty for reversed/garbage MRZ-like text without readable labels.
    if re.search(r"[A-Z0-9]{20,}", value) and signals == 0:
        score -= 20.0
    if signals == 0 and len(lines) > 20:
        score -= 10.0
    return score


def probe_orientation_text(image_path, engine="vietocr"):
    """Fast OCR pass used only to pick image rotation."""
    APP_STATE.ensure_models(engine)
    if engine == "paddle":
        _boxes, texts = APP_STATE.ocr(image_path, engine)
        return "\n".join(texts)

    from predict import predict_fast

    _boxes, texts = predict_fast(
        APP_STATE.detector,
        str(image_path),
        None,
        verbose=False,
        draw_output=False,
    )
    return "\n".join(texts)


def auto_orient_for_ocr(image_path, engine="vietocr"):
    """Rotate photo-style documents (CCCD/GCN shot sideways) to readable orientation."""
    import cv2
    from predict import imread_unicode, imwrite_unicode

    image = imread_unicode(str(image_path))
    if image is None:
        return image_path, 0, (0, 0)

    original_height, original_width = image.shape[:2]
    original_size = (original_width, original_height)
    is_portrait = original_height > original_width

    APP_STATE.ensure_models(engine)
    detector = APP_STATE.detector
    if detector is None:
        return image_path, 0, original_size

    rotate_flags = {
        0: None,
        90: cv2.ROTATE_90_CLOCKWISE,
        180: cv2.ROTATE_180,
        270: cv2.ROTATE_90_COUNTERCLOCKWISE,
    }

    ranked = []
    rotated_images = {}
    for angle, flag in rotate_flags.items():
        candidate = image if flag is None else cv2.rotate(image, flag)
        rotated_images[angle] = candidate
        score = score_detection_orientation(candidate, detector)
        height, width = candidate.shape[:2]
        if width >= height:
            score += 1.5
        ranked.append((score, angle))

    ranked.sort(reverse=True)

    # Portrait phone shots of landscape cards: detection aspect is more reliable than
    # paddle probe text (paddle can still read sideways lines; VietOCR cannot).
    if is_portrait:
        land_scores = {angle: score for score, angle in ranked if angle in (90, 270)}
        port_scores = {angle: score for score, angle in ranked if angle in (0, 180)}
        best_land_det = max(land_scores.values(), default=0.0)
        best_port_det = max(port_scores.values(), default=0.0)
        if best_land_det >= best_port_det + 12.0:
            finalist_angles = sorted(land_scores, key=land_scores.get, reverse=True)
        else:
            finalist_angles = [90, 270, 0]
            for _score, angle in ranked:
                if angle not in finalist_angles:
                    finalist_angles.append(angle)
                if len(finalist_angles) >= 4:
                    break
    else:
        # Landscape scans/PDF pages are usually already upright. Comparing 180° is
        # harmful: upside-down probe text can score "less bad" and flip page 1.
        finalist_angles = [0]
        for _score, angle in ranked:
            if angle == 180:
                continue
            if angle not in finalist_angles:
                finalist_angles.append(angle)
            if len(finalist_angles) >= 2:
                break

    best_angle = finalist_angles[0] if finalist_angles else 0
    best_content_score = None
    best_signals = -1
    for angle in finalist_angles:
        candidate = rotated_images[angle]
        height, width = candidate.shape[:2]
        temp_path = image_path.with_name(f"{image_path.stem}_orient_probe_{angle}{image_path.suffix}")
        imwrite_unicode(str(temp_path), candidate)
        try:
            joined = probe_orientation_text(temp_path, engine)
            content_score = score_ocr_text_quality(joined)
            signals = count_document_signals(joined)
            # Prefer orientations where detected text boxes are horizontal.
            det_score = next((score for score, cand_angle in ranked if cand_angle == angle), 0.0)
            content_score += 0.35 * det_score
            if is_portrait and width >= height:
                content_score += 8.0
            if has_strong_upright_cccd_signals(joined):
                content_score += 12.0
            # Land-certificate pages: reward readable field fragments even with OCR noise.
            if re.search(
                r"Th\S{0,3}a\s*đ|Di\S{0,2}n\s*t\S{0,2}ch|M[ụu]c\s*đ|Phường|Quận|Đất\s*ở|"
                r"S[ốô]\s*v[àa]o\s*s[ổô]|GIẤY\s*CHỨNG|t[ờò]\s*b[ảàáa]n\s*đ",
                joined,
                re.I,
            ):
                content_score += 18.0
        finally:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass

        better_signals = signals > best_signals
        better_score = best_content_score is None or content_score > best_content_score + 0.5
        similar_score = (
            best_content_score is not None
            and abs(content_score - best_content_score) <= 0.5
        )
        prefer_original = (
            similar_score
            and angle == 0
            and signals >= best_signals
            and (not is_portrait or has_strong_upright_cccd_signals(joined))
        )

        # Landscape scans are usually already upright; upside-down OCR can score
        # "less bad" on junk penalties, so require a clear win to leave angle 0.
        accept = better_signals or (signals == best_signals and better_score) or prefer_original
        if (
            not is_portrait
            and best_angle == 0
            and angle != 0
            and signals <= max(best_signals, 0)
        ):
            margin = 25.0 if signals == 0 else 12.0
            accept = content_score > (best_content_score or -999) + margin
        # Portrait: 180° flips are rarely correct for multi-page scans; need a clear win.
        if is_portrait and angle == 180 and best_angle != 180:
            margin = 20.0 if signals <= max(best_signals, 0) else 8.0
            accept = content_score > (best_content_score or -999) + margin and (
                signals > best_signals or has_strong_upright_cccd_signals(joined)
            )

        if accept:
            best_content_score = content_score
            best_signals = signals
            best_angle = angle

        # Fast path for already-landscape upright documents.
        if not is_portrait and angle == 0 and (
            has_strong_upright_cccd_signals(joined)
            or signals >= 2
            or re.search(
                r"Th\S{0,3}a\s*đ|Di\S{0,2}n\s*t\S{0,2}ch|M[ụu]c\s*đ|GIẤY\s*CHỨNG|Phường|Quận",
                joined,
                re.I,
            )
        ):
            best_angle = 0
            break

    if best_angle == 0:
        return image_path, 0, original_size

    oriented_path = image_path.with_name(f"{image_path.stem}_oriented{image_path.suffix}")
    imwrite_unicode(str(oriented_path), rotated_images[best_angle])
    return oriented_path, best_angle, original_size


def remap_point_to_original(x, y, angle, original_width, original_height):
    if angle == 90:
        return y, original_height - 1 - x
    if angle == 180:
        return original_width - 1 - x, original_height - 1 - y
    if angle == 270:
        return original_width - 1 - y, x
    return x, y


def remap_box_to_original(box, angle, original_size):
    if not angle or angle == 0 or not original_size:
        return box
    original_width, original_height = original_size
    remapped = []
    for point in serialize_box(box):
        x, y = remap_point_to_original(point[0], point[1], angle, original_width, original_height)
        remapped.append([x, y])
    return remapped


def prepare_ocr_image(image_path, engine="vietocr", auto_orient=True):
    import cv2
    from predict import imread_unicode, imwrite_unicode

    if auto_orient:
        oriented_path, angle, original_size = auto_orient_for_ocr(image_path, engine=engine)
    else:
        image = imread_unicode(str(image_path))
        if image is None:
            return image_path, 1.0, 0, (0, 0)
        height, width = image.shape[:2]
        oriented_path, angle, original_size = image_path, 0, (width, height)
    image = imread_unicode(str(oriented_path))
    if image is None:
        return oriented_path, 1.0, angle, original_size

    height, width = image.shape[:2]
    longest_side = max(width, height)
    # Small phone/scan pages need stronger upscale for VietOCR field extraction.
    upscale_cap = 2.6 if longest_side and longest_side < 1200 else OCR_UPSCALE_FACTOR
    scale = min(upscale_cap, OCR_MAX_SIDE / longest_side) if longest_side else 1.0
    if scale <= 1.05:
        return oriented_path, 1.0, angle, original_size

    resized = cv2.resize(
        image,
        None,
        fx=scale,
        fy=scale,
        interpolation=cv2.INTER_CUBIC,
    )
    ocr_path = oriented_path.with_name(f"{oriented_path.stem}_ocr{oriented_path.suffix}")
    imwrite_unicode(str(ocr_path), resized)
    return ocr_path, scale, angle, original_size


def scale_box(box, scale):
    if scale == 1.0:
        return box

    scaled = []
    for point in serialize_box(box):
        scaled.append([point[0] / scale, point[1] / scale])
    return scaled


def scale_lines(lines, scale):
    if scale == 1.0:
        return lines
    return [[value / scale for value in line] for line in lines]


def parse_multipart_form(body, content_type):
    match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type or "", re.I)
    if not match:
        raise ValueError("multipart/form-data thiếu boundary.")

    boundary = match.group("boundary").strip().strip('"').encode("utf-8")
    delimiter = b"--" + boundary
    fields = {}
    files = {}

    for part in body.split(delimiter):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        if part.startswith(b"--"):
            continue
        header_blob, _, content = part.partition(b"\r\n\r\n")
        if not _:
            header_blob, _, content = part.partition(b"\n\n")
        content = content.rstrip(b"\r\n")
        headers_text = header_blob.decode("utf-8", errors="replace")
        disposition = next(
            (line for line in headers_text.splitlines() if line.lower().startswith("content-disposition:")),
            "",
        )
        name_match = re.search(r'name="([^"]+)"', disposition, re.I)
        if not name_match:
            continue
        name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', disposition, re.I)
        if filename_match:
            upload = {
                "filename": Path(filename_match.group(1) or "upload.png").name,
                "content": content,
            }
            # Keep multiple uploads that share the same field name (e.g. two -F file=...).
            existing = files.get(name)
            if existing is None:
                files[name] = upload
            elif isinstance(existing, list):
                existing.append(upload)
            else:
                files[name] = [existing, upload]
        else:
            fields[name] = content.decode("utf-8", errors="replace")

    return fields, files


def flatten_multipart_uploads(files):
    uploads = []
    preferred_keys = [key for key in files.keys() if key.startswith("file")]
    other_keys = [key for key in files.keys() if key not in preferred_keys]
    for key in preferred_keys + other_keys:
        value = files[key]
        if isinstance(value, list):
            uploads.extend(item for item in value if item and item.get("content"))
        elif value and value.get("content"):
            uploads.append(value)
    return uploads


def normalize_upload_bytes(filename, file_bytes):
    suffix = Path(filename or "").suffix.lower()
    if suffix in PDF_EXTENSIONS:
        png_bytes = convert_pdf_bytes_to_png(file_bytes)
        return f"{Path(filename).stem or 'document'}.png", png_bytes
    if suffix and suffix not in IMAGE_EXTENSIONS:
        raise ValueError("Chỉ hỗ trợ ảnh hoặc PDF.")
    return filename or "upload.png", file_bytes


def save_upload_image(filename, file_bytes):
    filename, file_bytes = normalize_upload_bytes(filename, file_bytes)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    image_path = UPLOAD_DIR / safe_filename(filename)
    image_path.write_bytes(file_bytes)
    return image_path


def expand_upload_to_image_paths(upload):
    """Save one upload as one or more image paths (PDF → every page)."""
    filename = upload.get("filename") or "upload.png"
    content = upload.get("content") or b""
    suffix = Path(filename).suffix.lower()
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    if suffix in PDF_EXTENSIONS:
        paths = []
        for page_name, png_bytes in pdf_pages_as_named_pngs(filename, content):
            image_path = UPLOAD_DIR / safe_filename(page_name)
            image_path.write_bytes(png_bytes)
            paths.append(image_path)
        return paths
    return [save_upload_image(filename, content)]


def run_ocr_on_image(image_path, engine="vietocr", auto_orient=True):
    if engine not in {"vietocr", "paddle"}:
        raise ValueError("Engine không hợp lệ.")

    start_time = time.perf_counter()
    ocr_image_path, ocr_scale, angle, original_size = prepare_ocr_image(
        image_path,
        engine=engine,
        auto_orient=auto_orient,
    )
    boxes, texts = APP_STATE.ocr(ocr_image_path, engine)
    boxes = [scale_box(box, ocr_scale) for box in boxes]
    # Keep reading-order/table structure in the upright OCR orientation.
    table_rows = build_table_rows(boxes, texts)
    if angle:
        boxes = [remap_box_to_original(box, angle, original_size) for box in boxes]
    text = "\n".join(texts)
    elapsed_seconds = time.perf_counter() - start_time
    return {
        "text": text,
        "boxes": boxes,
        "texts": texts,
        "table_rows": table_rows,
        "ocr_scale": ocr_scale,
        "elapsed_seconds": elapsed_seconds,
    }


def export_document_json(image_path, json_type, engine="vietocr", auto_orient=True):
    """OCR one image and return the same JSON bytes as Export JSON (export_ocr.make_json)."""
    ocr_result = run_ocr_on_image(image_path, engine=engine, auto_orient=auto_orient)
    return make_json(ocr_result["text"], ocr_result["table_rows"], json_type)


def export_document_json_multi(image_paths, json_type, engine="vietocr", auto_orient=True):
    """OCR multiple pages then return one JSON (used by Giấy CNSDĐ Mẫu số 2)."""
    texts = []
    table_rows = []
    for image_path in image_paths:
        ocr_result = run_ocr_on_image(image_path, engine=engine, auto_orient=auto_orient)
        if ocr_result.get("text"):
            texts.append(ocr_result["text"])
        table_rows.extend(ocr_result.get("table_rows") or [])
    return make_json("\n".join(texts), table_rows, json_type)


class OCRRequestHandler(BaseHTTPRequestHandler):
    server_version = "VietnameseOCRWeb/1.0"

    def log_message(self, format, *args):
        return

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_bytes(INDEX_HTML.encode("utf-8"), "text/html; charset=utf-8")
            return

        if parsed.path == "/status":
            self.send_json(
                {
                    "engine": APP_STATE.current_engine,
                    "loading": APP_STATE.loading,
                    "error": APP_STATE.load_error,
                }
            )
            return

        if parsed.path == "/devices":
            self.handle_devices((parse_qs(parsed.query).get("kind") or [""])[0])
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path in DOCUMENT_API_ROUTES:
            self.handle_document_api(DOCUMENT_API_ROUTES[parsed.path])
            return

        if parsed.path == "/ocr":
            self.handle_ocr()
            return

        if parsed.path == "/export":
            query = parse_qs(parsed.query)
            self.handle_export((query.get("format") or ["docx"])[0])
            return

        if parsed.path == "/prepare-source":
            self.handle_prepare_source()
            return

        if parsed.path == "/capture-device":
            self.handle_capture_device()
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def read_upload_request(self):
        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)

        if "multipart/form-data" in content_type:
            fields, files = parse_multipart_form(body, content_type)
            uploads = flatten_multipart_uploads(files)
            if not uploads:
                raise ValueError("Thiếu file ảnh. Gửi field 'file' (Mẫu số 2: 2 file).")
            engine = (fields.get("engine") or "vietocr").strip() or "vietocr"
            return uploads, engine

        if "application/json" in content_type or body[:1] in (b"{", b"["):
            payload = json.loads(body.decode("utf-8"))
            engine = payload.get("engine", "vietocr")
            uploads = []
            if payload.get("images"):
                for index, item in enumerate(payload.get("images") or []):
                    filename = item.get("filename") or f"upload_{index + 1}.png"
                    image_data = item.get("image") or item.get("file") or ""
                    if image_data:
                        uploads.append({"filename": filename, "content": decode_data_url(image_data)})
            else:
                filename = payload.get("filename") or "upload.png"
                image_data = payload.get("image") or payload.get("file") or ""
                if image_data:
                    uploads.append({"filename": filename, "content": decode_data_url(image_data)})
            if not uploads:
                raise ValueError("Thiếu dữ liệu ảnh (image/file hoặc images[]).")
            return uploads, engine

        raise ValueError("Hỗ trợ multipart/form-data (file) hoặc application/json (image base64).")

    def handle_document_api(self, json_type):
        try:
            uploads, engine = self.read_upload_request()
            image_paths = []
            from_pdf = False
            for upload in uploads:
                filename = upload.get("filename") or ""
                if Path(filename).suffix.lower() in PDF_EXTENSIONS:
                    from_pdf = True
                image_paths.extend(expand_upload_to_image_paths(upload))
            if not image_paths:
                raise ValueError("Thiếu file ảnh hoặc PDF.")

            # PDF pages from PyMuPDF are already upright.
            auto_orient = not from_pdf

            if json_type == "land_certificate_2":
                if len(image_paths) < 2:
                    raise ValueError(
                        "Mẫu số 2 cần ít nhất 2 trang. Gửi 2 ảnh, hoặc 1 PDF có ≥ 2 trang."
                    )
                # OCR và gộp mọi trang (PDF nhiều trang hoặc nhiều ảnh).
                data = export_document_json_multi(
                    image_paths,
                    json_type,
                    engine=engine,
                    auto_orient=auto_orient,
                )
            else:
                data = export_document_json(
                    image_paths[0],
                    json_type,
                    engine=engine,
                    auto_orient=auto_orient,
                )
            self.send_bytes(data, "application/json; charset=utf-8")
        except ValueError as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_devices(self, kind):
        try:
            if kind == "printer":
                devices = list_windows_printers()
            elif kind == "scanner":
                devices = list_windows_scanners()
            else:
                raise ValueError("Loại thiết bị không hợp lệ.")
            self.send_json({"ok": True, "kind": kind, "devices": devices})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_prepare_source(self):
        try:
            payload = self.read_json()
            filename = payload.get("filename") or "upload.bin"
            file_bytes = decode_data_url(payload.get("file", ""))
            suffix = Path(filename).suffix.lower()

            if suffix in PDF_EXTENSIONS or (payload.get("file") or "").startswith("data:application/pdf"):
                named_pages = pdf_pages_as_named_pngs(filename, file_bytes)
                pages = [
                    {
                        "filename": page_name,
                        "image": image_bytes_to_data_url(png_bytes),
                    }
                    for page_name, png_bytes in named_pages
                ]
                self.send_json(
                    {
                        "ok": True,
                        "filename": pages[0]["filename"],
                        "image": pages[0]["image"],
                        "page_count": len(pages),
                        "pages": pages,
                    }
                )
                return

            if suffix not in IMAGE_EXTENSIONS:
                raise ValueError("Chỉ hỗ trợ ảnh hoặc PDF.")

            mime = "image/png"
            if suffix in {".jpg", ".jpeg"}:
                mime = "image/jpeg"
            elif suffix == ".webp":
                mime = "image/webp"
            elif suffix == ".bmp":
                mime = "image/bmp"
            elif suffix in {".tif", ".tiff"}:
                mime = "image/tiff"

            self.send_json(
                {
                    "ok": True,
                    "filename": Path(filename).name,
                    "image": image_bytes_to_data_url(file_bytes, mime),
                }
            )
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def handle_capture_device(self):
        try:
            payload = self.read_json()
            kind = payload.get("kind")
            device_id = payload.get("device_id") or ""
            device_name = payload.get("device_name") or device_id or kind

            if kind == "scanner":
                image_bytes = capture_scanner_image(device_id)
                filename = f"{Path(str(device_name)).stem or 'scanner'}_scan.png"
            elif kind == "printer":
                raise ValueError(
                    "Đã kết nối máy in, nhưng chưa hỗ trợ lấy ảnh trực tiếp từ máy in. "
                    "Hãy dùng 'Chọn ảnh/pdf' hoặc 'Kết nối máy scan'."
                )
            else:
                raise ValueError("Loại thiết bị không hợp lệ.")

            self.send_json(
                {
                    "ok": True,
                    "filename": filename,
                    "image": image_bytes_to_data_url(image_bytes),
                }
            )
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length)
        return json.loads(raw_body.decode("utf-8"))

    def handle_ocr(self):
        try:
            payload = self.read_json()
            engine = payload.get("engine", "vietocr")
            if engine not in {"vietocr", "paddle"}:
                raise ValueError("Engine không hợp lệ.")

            image_data = payload.get("image", "")
            if "," in image_data:
                _header, image_data = image_data.split(",", 1)

            image_bytes = base64.b64decode(image_data)
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            image_path = UPLOAD_DIR / safe_filename(payload.get("filename"))
            image_path.write_bytes(image_bytes)

            start_time = time.perf_counter()
            auto_orient = payload.get("auto_orient", True)
            if isinstance(auto_orient, str):
                auto_orient = auto_orient.strip().lower() not in {"0", "false", "no", "off"}
            ocr_image_path, ocr_scale, angle, original_size = prepare_ocr_image(
                image_path,
                engine=engine,
                auto_orient=bool(auto_orient),
            )
            boxes, texts = APP_STATE.ocr(ocr_image_path, engine)
            boxes = [scale_box(box, ocr_scale) for box in boxes]
            # Keep layout/export in upright (readable) orientation, same as landscape shots.
            display_path = image_path
            if angle:
                oriented_candidate = image_path.with_name(f"{image_path.stem}_oriented{image_path.suffix}")
                if oriented_candidate.exists():
                    display_path = oriented_candidate

            table_rows = build_table_rows(boxes, texts)
            table_lines = scale_lines(detect_table_lines(ocr_image_path), ocr_scale)
            text = "\n".join(texts)
            serialized_boxes = [serialize_box(box) for box in boxes]
            styles = estimate_text_styles(display_path, boxes, texts)
            items = [
                {"text": str(text_value), "box": box, "style": styles.get(index, {})}
                for index, (text_value, box) in enumerate(zip(texts, serialized_boxes))
                if str(text_value).strip() and box
            ]
            table_cells, table_indices, table_grid = build_table_cells(boxes, texts, table_lines, styles)
            layout_rows = build_layout_rows(boxes, texts, table_indices, styles)
            image_width, image_height = get_image_size(display_path)
            elapsed_seconds = time.perf_counter() - start_time

            response_payload = {
                "ok": True,
                "text": text,
                "boxes": serialized_boxes,
                "items": items,
                "layout_rows": layout_rows,
                "table_cells": table_cells,
                "table_grid": table_grid,
                "image_width": image_width,
                "image_height": image_height,
                "table_rows": table_rows,
                "table_lines": table_lines,
                "ocr_scale": ocr_scale,
                "orientation_angle": angle,
                "elapsed_seconds": elapsed_seconds,
            }
            if angle and display_path != image_path:
                suffix = display_path.suffix.lower()
                mime = "image/jpeg" if suffix in {".jpg", ".jpeg"} else "image/png"
                response_payload["oriented_image"] = image_bytes_to_data_url(display_path.read_bytes(), mime)

            self.send_json(response_payload)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_export(self, file_format):
        try:
            payload = self.read_json()
            pages = payload.get("pages") or []
            text = payload.get("text", "").strip()
            if not text and not any(str(page.get("text") or "").strip() for page in pages):
                raise ValueError("Chưa có kết quả OCR để export.")

            if file_format == "xlsx":
                table_rows = payload.get("table_rows") or None
                data = make_xlsx(text, table_rows, pages=pages or None)
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                filename = payload.get("filename") or "ocr_result.xlsx"
            elif file_format == "docx":
                data = make_docx(text, payload.get("formatted_data") or {}, pages=pages or None)
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                filename = payload.get("filename") or "ocr_result.docx"
            elif file_format == "json":
                table_rows = payload.get("table_rows") or None
                data = make_json(text, table_rows, payload.get("json_type") or "cccd_front")
                content_type = "application/json; charset=utf-8"
                filename = payload.get("filename") or "ocr_result.json"
            else:
                raise ValueError("Định dạng export không hỗ trợ.")

            self.send_bytes(
                data,
                content_type,
                headers={
                    "Content-Disposition": f'attachment; filename="{Path(filename).name}"',
                },
            )
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def send_json(self, data, status=HTTPStatus.OK):
        self.send_bytes(
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            "application/json; charset=utf-8",
            status,
        )

    def send_bytes(self, data, content_type=None, status=HTTPStatus.OK, headers=None):
        self.send_response(status)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Type", content_type or "application/octet-stream")
        self.send_header("Cache-Control", "no-store")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)


def main():
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    server = ThreadingHTTPServer((HOST, PORT), OCRRequestHandler)
    local_ip = get_local_ip()
    local_url = f"http://127.0.0.1:{PORT}"
    lan_url = f"http://{local_ip}:{PORT}"
    print(f"Vietnamese OCR web UI local: {local_url}", flush=True)
    print(f"Vietnamese OCR web UI LAN:   {lan_url}", flush=True)
    print("Document APIs:", flush=True)
    print(f"  POST {local_url}/api/cccd-front", flush=True)
    print(f"  POST {local_url}/api/cccd-back", flush=True)
    print(f"  POST {local_url}/api/land-certificate", flush=True)
    print(f"  POST {local_url}/api/land-certificate-2", flush=True)
    print(f"  POST {local_url}/api/land-certificate-3", flush=True)
    APP_STATE.preload("vietocr")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
