const state = {
  rows: [],
  headers: [],
  profile: [],
  report: null,
  fileName: "demo-dataset.csv",
};

const els = {
  fileInput: document.getElementById("fileInput"),
  dropZone: document.getElementById("dropZone"),
  fileMeta: document.getElementById("fileMeta"),
  statusText: document.getElementById("statusText"),
  scanTime: document.getElementById("scanTime"),
  statusDot: document.querySelector(".status-dot"),
  qualityScore: document.getElementById("qualityScore"),
  qualityLabel: document.getElementById("qualityLabel"),
  rowCount: document.getElementById("rowCount"),
  colCount: document.getElementById("colCount"),
  missingCount: document.getElementById("missingCount"),
  missingPct: document.getElementById("missingPct"),
  duplicateCount: document.getElementById("duplicateCount"),
  outlierCount: document.getElementById("outlierCount"),
  profileBody: document.getElementById("profileBody"),
  columnSearch: document.getElementById("columnSearch"),
  targetSelect: document.getElementById("targetSelect"),
  balanceResult: document.getElementById("balanceResult"),
  balanceBars: document.getElementById("balanceBars"),
  issueList: document.getElementById("issueList"),
  previewTable: document.getElementById("previewTable"),
  loadDemoBtn: document.getElementById("loadDemoBtn"),
  resetBtn: document.getElementById("resetBtn"),
  downloadReportBtn: document.getElementById("downloadReportBtn"),
  toast: document.getElementById("toast"),
};

const missingTokens = new Set(["", "na", "n/a", "nan", "null", "none", "undefined", "-", "--", "?"]);

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => els.toast.classList.remove("show"), 2600);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatNumber(value) {
  if (Number.isNaN(value) || value === null || value === undefined) return "--";
  return new Intl.NumberFormat().format(value);
}

function formatPct(value, digits = 1) {
  if (!Number.isFinite(value)) return "--";
  return `${value.toFixed(digits)}%`;
}

function isMissing(value) {
  if (value === null || value === undefined) return true;
  const clean = String(value).trim().toLowerCase();
  return missingTokens.has(clean);
}

function normalizeCell(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).trim();
}

function toNumber(value) {
  if (isMissing(value)) return null;
  const clean = String(value).replace(/,/g, "").trim();
  if (clean === "") return null;
  const num = Number(clean);
  return Number.isFinite(num) ? num : null;
}

function quantile(sortedValues, q) {
  if (!sortedValues.length) return null;
  const position = (sortedValues.length - 1) * q;
  const base = Math.floor(position);
  const rest = position - base;
  if (sortedValues[base + 1] !== undefined) {
    return sortedValues[base] + rest * (sortedValues[base + 1] - sortedValues[base]);
  }
  return sortedValues[base];
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quote = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (quote && next === '"') {
        cell += '"';
        i += 1;
      } else {
        quote = !quote;
      }
    } else if (char === "," && !quote) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quote) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell);
      if (row.some(value => String(value).trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell);
  if (row.some(value => String(value).trim() !== "")) rows.push(row);

  if (!rows.length) return { headers: [], dataRows: [] };

  let headers = rows[0].map((h, index) => {
    const name = normalizeCell(h) || `Column_${index + 1}`;
    return name;
  });
  headers = makeUniqueHeaders(headers);

  const dataRows = rows.slice(1).map(rawRow => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = normalizeCell(rawRow[index]);
    });
    return record;
  });

  return { headers, dataRows };
}

function makeUniqueHeaders(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = header || `Column_${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function parseWorkbook(arrayBuffer) {
  if (!window.XLSX) {
    throw new Error("XLSX parser is not loaded. Connect to the internet or upload CSV instead.");
  }
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const firstSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheet];
  const raw = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: false });
  if (!raw.length) return { headers: [], dataRows: [] };
  const headers = makeUniqueHeaders(raw[0].map((h, index) => normalizeCell(h) || `Column_${index + 1}`));
  const dataRows = raw.slice(1).map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = normalizeCell(row[index]);
    });
    return record;
  });
  return { headers, dataRows };
}

async function handleFile(file) {
  if (!file) return;
  state.fileName = file.name;
  setStatus("Scanning dataset...", "Reading file locally.", false);
  els.fileMeta.textContent = `${file.name} • ${(file.size / 1024).toFixed(1)} KB`;

  const extension = file.name.split(".").pop().toLowerCase();
  try {
    let parsed;
    if (["xlsx", "xls"].includes(extension)) {
      const buffer = await file.arrayBuffer();
      parsed = parseWorkbook(buffer);
    } else {
      const text = await file.text();
      parsed = parseCsv(text);
    }

    if (!parsed.headers.length || !parsed.dataRows.length) {
      throw new Error("No readable rows found. Check that your file has headers and data.");
    }

    scanDataset(parsed.headers, parsed.dataRows);
  } catch (error) {
    console.error(error);
    setStatus("Scan failed", error.message, false);
    showToast(error.message);
  }
}

function scanDataset(headers, rows) {
  const start = performance.now();
  state.headers = headers;
  state.rows = rows;

  const duplicateInfo = detectDuplicates(rows, headers);
  const profile = headers.map(header => profileColumn(header, rows));
  const totalMissing = profile.reduce((sum, col) => sum + col.missingCount, 0);
  const totalOutliers = profile.reduce((sum, col) => sum + col.outlierCount, 0);
  const totalCells = rows.length * headers.length;
  const missingPct = totalCells ? (totalMissing / totalCells) * 100 : 0;
  const duplicatePct = rows.length ? (duplicateInfo.duplicateCount / rows.length) * 100 : 0;
  const outlierPct = totalCells ? (totalOutliers / totalCells) * 100 : 0;
  const riskyColumnCount = profile.filter(col => col.risk !== "Low").length;
  const qualityScore = calculateQualityScore({ missingPct, duplicatePct, outlierPct, riskyColumnCount, colCount: headers.length });

  const issues = buildIssues({ profile, rows, headers, duplicateInfo, missingPct, duplicatePct, totalOutliers, qualityScore });

  state.profile = profile;
  state.report = {
    fileName: state.fileName,
    generatedAt: new Date().toISOString(),
    summary: {
      rows: rows.length,
      columns: headers.length,
      missingCells: totalMissing,
      missingPercent: Number(missingPct.toFixed(3)),
      duplicateRows: duplicateInfo.duplicateCount,
      duplicatePercent: Number(duplicatePct.toFixed(3)),
      outlierFlags: totalOutliers,
      outlierPercent: Number(outlierPct.toFixed(3)),
      qualityScore,
      qualityLabel: qualityLabel(qualityScore),
    },
    columns: profile,
    issues,
  };

  updateSummary(state.report.summary);
  renderProfile(profile);
  renderTargetOptions(profile);
  renderIssues(issues);
  renderPreview(headers, rows);

  const elapsed = Math.max(1, Math.round(performance.now() - start));
  setStatus("Dataset scan complete", `${state.fileName} analyzed in ${elapsed} ms.`, true);
  els.downloadReportBtn.disabled = false;
  showToast("Dataset quality scan complete.");
}

function profileColumn(header, rows) {
  const values = rows.map(row => row[header]);
  const nonMissing = values.filter(value => !isMissing(value));
  const missingCount = values.length - nonMissing.length;
  const uniqueValues = new Set(nonMissing.map(value => String(value).trim()));
  const numericValues = nonMissing.map(toNumber).filter(value => value !== null);
  const numericRatio = nonMissing.length ? numericValues.length / nonMissing.length : 0;
  const type = inferType(nonMissing, numericValues, numericRatio);
  const outlierInfo = type === "Numeric" ? detectOutliers(numericValues) : { count: 0, lower: null, upper: null };
  const missingPct = values.length ? (missingCount / values.length) * 100 : 0;
  const uniquePct = values.length ? (uniqueValues.size / values.length) * 100 : 0;

  let risk = "Low";
  const notes = [];
  if (missingPct >= 30) {
    risk = "High";
    notes.push("Large missing-value share");
  } else if (missingPct >= 10) {
    risk = "Medium";
    notes.push("Moderate missing-value share");
  }
  if (outlierInfo.count >= Math.max(5, numericValues.length * 0.08)) {
    risk = risk === "High" ? "High" : "Medium";
    notes.push("Many numeric outliers");
  }
  if (uniqueValues.size === 1 && nonMissing.length > 0) {
    risk = "Medium";
    notes.push("Constant column");
  }
  if (type === "Mixed") {
    risk = risk === "High" ? "High" : "Medium";
    notes.push("Mixed numeric/text values");
  }

  return {
    name: header,
    type,
    rows: values.length,
    missingCount,
    missingPercent: Number(missingPct.toFixed(2)),
    uniqueCount: uniqueValues.size,
    uniquePercent: Number(uniquePct.toFixed(2)),
    outlierCount: outlierInfo.count,
    outlierLowerBound: outlierInfo.lower,
    outlierUpperBound: outlierInfo.upper,
    min: type === "Numeric" && numericValues.length ? Math.min(...numericValues) : null,
    max: type === "Numeric" && numericValues.length ? Math.max(...numericValues) : null,
    risk,
    notes,
    sampleValues: Array.from(uniqueValues).slice(0, 5),
  };
}

function inferType(nonMissing, numericValues, numericRatio) {
  if (nonMissing.length === 0) return "Empty";
  if (numericRatio >= 0.95) return "Numeric";
  if (numericRatio > 0.05 && numericRatio < 0.95) return "Mixed";

  const dateLikeCount = nonMissing.filter(value => {
    const clean = String(value).trim();
    if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(clean)) return true;
    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(clean)) return true;
    const time = Date.parse(clean);
    return Number.isFinite(time) && /[-/:]/.test(clean);
  }).length;
  if (nonMissing.length && dateLikeCount / nonMissing.length >= 0.8) return "Date";

  const uniqueCount = new Set(nonMissing.map(value => String(value).trim())).size;
  if (uniqueCount <= Math.max(20, Math.sqrt(nonMissing.length) * 2)) return "Categorical";
  return "Text";
}

function detectOutliers(numericValues) {
  if (numericValues.length < 4) return { count: 0, lower: null, upper: null };
  const sorted = [...numericValues].sort((a, b) => a - b);
  const q1 = quantile(sorted, 0.25);
  const q3 = quantile(sorted, 0.75);
  const iqr = q3 - q1;
  if (!Number.isFinite(iqr) || iqr === 0) return { count: 0, lower: q1, upper: q3 };
  const lower = q1 - 1.5 * iqr;
  const upper = q3 + 1.5 * iqr;
  const count = numericValues.filter(value => value < lower || value > upper).length;
  return { count, lower: Number(lower.toFixed(4)), upper: Number(upper.toFixed(4)) };
}

function detectDuplicates(rows, headers) {
  const seen = new Set();
  let duplicateCount = 0;
  rows.forEach(row => {
    const key = headers.map(header => normalizeCell(row[header]).toLowerCase()).join("␟");
    if (seen.has(key)) duplicateCount += 1;
    else seen.add(key);
  });
  return { duplicateCount, uniqueRows: seen.size };
}

function calculateQualityScore({ missingPct, duplicatePct, outlierPct, riskyColumnCount, colCount }) {
  let score = 100;
  score -= Math.min(35, missingPct * 1.4);
  score -= Math.min(20, duplicatePct * 1.2);
  score -= Math.min(20, outlierPct * 2.2);
  score -= Math.min(20, colCount ? (riskyColumnCount / colCount) * 20 : 0);
  return Math.max(0, Math.round(score));
}

function qualityLabel(score) {
  if (score >= 85) return "Excellent";
  if (score >= 70) return "Good";
  if (score >= 50) return "Needs cleaning";
  return "High risk";
}

function riskBadge(risk) {
  if (risk === "High") return `<span class="badge bad">High</span>`;
  if (risk === "Medium") return `<span class="badge warn">Medium</span>`;
  return `<span class="badge good">Low</span>`;
}

function typeBadge(type) {
  const badgeType = type === "Numeric" ? "info" : type === "Mixed" ? "warn" : type === "Empty" ? "bad" : "good";
  return `<span class="badge ${badgeType}">${escapeHtml(type)}</span>`;
}

function updateSummary(summary) {
  els.qualityScore.textContent = summary.qualityScore;
  els.qualityLabel.textContent = summary.qualityLabel;
  els.rowCount.textContent = formatNumber(summary.rows);
  els.colCount.textContent = formatNumber(summary.columns);
  els.missingCount.textContent = formatNumber(summary.missingCells);
  els.missingPct.textContent = `${formatPct(summary.missingPercent)} of all cells`;
  els.duplicateCount.textContent = formatNumber(summary.duplicateRows);
  els.outlierCount.textContent = formatNumber(summary.outlierFlags);
}

function renderProfile(profile) {
  const query = els.columnSearch.value.trim().toLowerCase();
  const filtered = profile.filter(col => col.name.toLowerCase().includes(query));
  if (!filtered.length) {
    els.profileBody.innerHTML = `<tr><td colspan="6" class="empty-cell">No matching columns found.</td></tr>`;
    return;
  }

  els.profileBody.innerHTML = filtered.map(col => `
    <tr title="${escapeHtml(col.notes.join(" • ") || "No major issue detected")}">
      <td><strong>${escapeHtml(col.name)}</strong><br><small>${escapeHtml(col.sampleValues.join(", ") || "No sample")}</small></td>
      <td>${typeBadge(col.type)}</td>
      <td>${formatNumber(col.missingCount)} <small>(${formatPct(col.missingPercent)})</small></td>
      <td>${formatNumber(col.uniqueCount)} <small>(${formatPct(col.uniquePercent)})</small></td>
      <td>${formatNumber(col.outlierCount)}</td>
      <td>${riskBadge(col.risk)}</td>
    </tr>
  `).join("");
}

function renderTargetOptions(profile) {
  const candidates = profile.filter(col => ["Categorical", "Text", "Mixed"].includes(col.type) && col.uniqueCount > 1 && col.uniqueCount <= 50);
  els.targetSelect.disabled = candidates.length === 0;
  if (!candidates.length) {
    els.targetSelect.innerHTML = `<option>No categorical target found</option>`;
    els.balanceResult.innerHTML = `<span>No suitable target column found.</span>`;
    els.balanceBars.innerHTML = "";
    return;
  }

  els.targetSelect.innerHTML = candidates.map(col => `<option value="${escapeHtml(col.name)}">${escapeHtml(col.name)}</option>`).join("");
  renderBalance(candidates[0].name);
}

function renderBalance(columnName) {
  if (!columnName || !state.rows.length) return;
  const counts = new Map();
  state.rows.forEach(row => {
    const value = isMissing(row[columnName]) ? "[Missing]" : normalizeCell(row[columnName]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });

  const sorted = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  const total = state.rows.length;
  const majority = sorted[0]?.[1] || 0;
  const minority = sorted.filter(([, count]) => count > 0).at(-1)?.[1] || 0;
  const ratio = minority ? majority / minority : Infinity;
  const imbalance = ratio >= 5 || majority / total >= 0.85;

  els.balanceResult.innerHTML = `
    <span>${escapeHtml(columnName)} has <strong>${sorted.length}</strong> classes</span>
    <span class="badge ${imbalance ? "bad" : ratio >= 2.5 ? "warn" : "good"}">${imbalance ? "Imbalanced" : ratio >= 2.5 ? "Mild skew" : "Balanced"}</span>
  `;

  const max = majority || 1;
  els.balanceBars.innerHTML = sorted.slice(0, 12).map(([label, count]) => `
    <div class="bar-row">
      <div class="bar-meta"><span>${escapeHtml(label)}</span><span>${formatNumber(count)} • ${formatPct((count / total) * 100)}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(3, (count / max) * 100)}%"></div></div>
    </div>
  `).join("");
}

function buildIssues({ profile, rows, headers, duplicateInfo, missingPct, duplicatePct, totalOutliers, qualityScore }) {
  const issues = [];
  const highMissing = profile.filter(col => col.missingPercent >= 20).sort((a, b) => b.missingPercent - a.missingPercent).slice(0, 4);
  const mixedColumns = profile.filter(col => col.type === "Mixed");
  const outlierColumns = profile.filter(col => col.outlierCount > 0).sort((a, b) => b.outlierCount - a.outlierCount).slice(0, 4);
  const constantColumns = profile.filter(col => col.uniqueCount === 1 && col.missingCount < rows.length);
  const emptyColumns = profile.filter(col => col.type === "Empty");

  if (missingPct > 0) {
    issues.push({
      icon: "∅",
      title: `${formatPct(missingPct)} total missing values`,
      text: highMissing.length
        ? `Highest missing columns: ${highMissing.map(col => `${col.name} (${formatPct(col.missingPercent)})`).join(", ")}. Consider imputation, removal, or source correction.`
        : "Missing values exist across the dataset. Decide whether to impute, remove, or flag them before modeling.",
    });
  }

  if (duplicateInfo.duplicateCount > 0) {
    issues.push({
      icon: "≡",
      title: `${formatNumber(duplicateInfo.duplicateCount)} duplicate rows detected`,
      text: `That is ${formatPct(duplicatePct)} of all rows. Review whether duplicates are valid repeated events or accidental copies.`,
    });
  }

  if (totalOutliers > 0) {
    issues.push({
      icon: "↯",
      title: `${formatNumber(totalOutliers)} numeric outlier flags`,
      text: `Most affected: ${outlierColumns.map(col => `${col.name} (${col.outlierCount})`).join(", ")}. Validate extreme values before training or reporting.`,
    });
  }

  if (mixedColumns.length) {
    issues.push({
      icon: "T#",
      title: `${mixedColumns.length} mixed-type columns`,
      text: `${mixedColumns.map(col => col.name).slice(0, 5).join(", ")} contain both numeric-like and text-like values. Standardize formats before analysis.`,
    });
  }

  if (constantColumns.length || emptyColumns.length) {
    issues.push({
      icon: "—",
      title: `${constantColumns.length + emptyColumns.length} low-information columns`,
      text: `Constant or empty columns rarely help models. Review: ${[...constantColumns, ...emptyColumns].map(col => col.name).slice(0, 6).join(", ")}.`,
    });
  }

  if (qualityScore >= 85 && issues.length === 0) {
    issues.push({ icon: "✓", title: "Dataset looks clean", text: "No major missingness, duplicates, or outlier risks were detected by this quick scanner." });
  } else if (!issues.length) {
    issues.push({ icon: "i", title: "No critical issue found", text: "The quick checks passed, but you should still validate domain-specific rules and label correctness." });
  }

  return issues;
}

function renderIssues(issues) {
  els.issueList.innerHTML = issues.map(issue => `
    <div class="issue-item">
      <div class="issue-icon">${escapeHtml(issue.icon)}</div>
      <div><strong>${escapeHtml(issue.title)}</strong><p>${escapeHtml(issue.text)}</p></div>
    </div>
  `).join("");
}

function renderPreview(headers, rows) {
  const previewRows = rows.slice(0, 20);
  els.previewTable.innerHTML = `
    <thead><tr>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>
      ${previewRows.map(row => `<tr>${headers.map(header => `<td>${escapeHtml(row[header])}</td>`).join("")}</tr>`).join("")}
    </tbody>
  `;
}

function setStatus(title, subtitle, ready) {
  els.statusText.textContent = title;
  els.scanTime.textContent = subtitle;
  els.statusDot.classList.toggle("ready", ready);
}

function downloadReport() {
  if (!state.report) return;
  const blob = new Blob([JSON.stringify(state.report, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const safeName = state.fileName.replace(/\.[^.]+$/, "").replace(/[^a-z0-9_-]/gi, "_").toLowerCase();
  link.href = url;
  link.download = `${safeName || "dataset"}_quality_report.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showToast("Quality report downloaded.");
}

function resetApp() {
  state.rows = [];
  state.headers = [];
  state.profile = [];
  state.report = null;
  state.fileName = "demo-dataset.csv";
  els.fileInput.value = "";
  els.fileMeta.textContent = "No file selected";
  els.qualityScore.textContent = "--";
  els.qualityLabel.textContent = "No scan yet";
  els.rowCount.textContent = "--";
  els.colCount.textContent = "--";
  els.missingCount.textContent = "--";
  els.missingPct.textContent = "--";
  els.duplicateCount.textContent = "--";
  els.outlierCount.textContent = "--";
  els.profileBody.innerHTML = `<tr><td colspan="6" class="empty-cell">Upload data to view profiling results.</td></tr>`;
  els.previewTable.innerHTML = `<tbody><tr><td class="empty-cell">Dataset preview will appear here.</td></tr></tbody>`;
  els.targetSelect.disabled = true;
  els.targetSelect.innerHTML = `<option>No dataset loaded</option>`;
  els.balanceResult.innerHTML = `<span>Choose a target column after upload.</span>`;
  els.balanceBars.innerHTML = "";
  els.issueList.innerHTML = `<div class="empty-state">No issues yet. Run a scan first.</div>`;
  els.downloadReportBtn.disabled = true;
  setStatus("Waiting for a dataset", "Upload a file or use demo data to begin.", false);
  showToast("Scanner reset.");
}

function loadDemo() {
  const headers = ["customer_id", "age", "city", "plan", "monthly_spend", "churn", "signup_date", "support_calls"];
  const rows = [
    ["C001", "23", "Dhaka", "Basic", "290", "No", "2025-01-05", "1"],
    ["C002", "42", "Chittagong", "Premium", "890", "Yes", "2025-01-07", "5"],
    ["C003", "", "Khulna", "Basic", "310", "No", "2025-01-08", "0"],
    ["C004", "35", "Dhaka", "Standard", "520", "No", "2025-01-09", "2"],
    ["C005", "64", "Sylhet", "Premium", "910", "Yes", "2025-01-11", "8"],
    ["C006", "28", "Dhaka", "Basic", "280", "No", "2025-01-12", "1"],
    ["C007", "31", "Rajshahi", "Standard", "560", "No", "2025-01-12", "2"],
    ["C008", "29", "Dhaka", "Basic", "", "No", "2025-01-14", "1"],
    ["C009", "51", "Khulna", "Premium", "960", "Yes", "2025-01-16", "9"],
    ["C010", "27", "Dhaka", "Basic", "305", "No", "2025-01-18", "0"],
    ["C011", "25", "Chittagong", "Basic", "315", "No", "2025-01-19", "1"],
    ["C012", "46", "Dhaka", "Standard", "540", "No", "2025-01-20", "3"],
    ["C013", "37", "Barisal", "Standard", "570", "No", "2025-01-21", "2"],
    ["C014", "34", "Dhaka", "Basic", "295", "No", "2025-01-22", "1"],
    ["C015", "45", "Sylhet", "Premium", "880", "Yes", "2025-01-23", "7"],
    ["C016", "200", "Dhaka", "Basic", "2500", "No", "2025-01-25", "0"],
    ["C017", "33", "Comilla", "Standard", "bad", "No", "2025-01-26", "2"],
    ["C018", "41", "Dhaka", "Basic", "300", "No", "2025-01-27", "1"],
    ["C018", "41", "Dhaka", "Basic", "300", "No", "2025-01-27", "1"],
    ["C020", "58", "Chittagong", "Premium", "930", "Yes", "2025-01-29", "10"],
    ["C021", "30", "Dhaka", "Basic", "285", "No", "2025-01-31", "0"],
    ["C022", "", "Dhaka", "Basic", "305", "No", "2025-02-02", "1"],
    ["C023", "39", "Khulna", "Standard", "550", "No", "2025-02-03", "2"],
    ["C024", "48", "Dhaka", "Premium", "900", "Yes", "2025-02-04", "6"],
    ["C025", "26", "Dhaka", "Basic", "299", "No", "2025-02-05", "1"],
  ].map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
  state.fileName = "demo_telco_churn.csv";
  els.fileMeta.textContent = "demo_telco_churn.csv • sample data";
  scanDataset(headers, rows);
}

els.fileInput.addEventListener("change", event => handleFile(event.target.files[0]));
els.loadDemoBtn.addEventListener("click", loadDemo);
els.resetBtn.addEventListener("click", resetApp);
els.downloadReportBtn.addEventListener("click", downloadReport);
els.columnSearch.addEventListener("input", () => renderProfile(state.profile));
els.targetSelect.addEventListener("change", event => renderBalance(event.target.value));

["dragenter", "dragover"].forEach(eventName => {
  els.dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    els.dropZone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach(eventName => {
  els.dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    els.dropZone.classList.remove("dragover");
  });
});

els.dropZone.addEventListener("drop", event => {
  const file = event.dataTransfer.files[0];
  handleFile(file);
});

els.dropZone.addEventListener("click", () => els.fileInput.click());
