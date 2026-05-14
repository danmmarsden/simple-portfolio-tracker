const STORAGE_KEY = "share-tracker-state-v2";
const LEGACY_STORAGE_KEY = "share-tracker-state-v1";
const DEFAULT_USD_TO_GBP_RATE = 0.79;

const currencyFormatter = new Intl.NumberFormat("en-GB", {
  style: "currency",
  currency: "GBP",
});

const percentFormatter = new Intl.NumberFormat("en-GB", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const elements = {
  sheetSettingsForm: document.getElementById("sheet-settings-form"),
  sheetCsvUrlInput: document.getElementById("sheet-csv-url-input"),
  sheetStatus: document.getElementById("sheet-status"),
  toggleSheetSync: document.getElementById("toggle-sheet-sync"),
  sheetSyncContent: document.getElementById("sheet-sync-content"),
  sheetSyncToggleLabel: document.getElementById("sheet-sync-toggle-label"),
  syncSheetNow: document.getElementById("sync-sheet-now"),
  clearSheetSource: document.getElementById("clear-sheet-source"),
  holdingsSection: document.getElementById("holdings-section"),
  tableBody: document.getElementById("holdings-table-body"),
  tableFoot: document.getElementById("holdings-table-foot"),
  holdingsLastUpdated: document.getElementById("holdings-last-updated"),
  emptyState: document.getElementById("empty-state"),
  summaryDate: document.getElementById("summary-date"),
  summaryChart: document.getElementById("summary-chart"),
  summaryChartRange: document.getElementById("summary-chart-range"),
  summaryChartEmpty: document.getElementById("summary-chart-empty"),
  totalValue: document.getElementById("total-value"),
  totalChange: document.getElementById("total-change"),
  totalChangePercent: document.getElementById("total-change-percent"),
  refreshAll: document.getElementById("refresh-all"),
};

const state = loadState();

elements.sheetSettingsForm.addEventListener("submit", handleSheetSettingsSubmit);
elements.tableBody.addEventListener("click", handleTableClick);
elements.tableBody.addEventListener("keydown", handleTableKeydown);
elements.refreshAll.addEventListener("click", refreshAllQuotes);
elements.syncSheetNow.addEventListener("click", syncSheetDataNow);
elements.clearSheetSource.addEventListener("click", clearSheetSource);
elements.toggleSheetSync.addEventListener("click", toggleSheetSyncPanel);

elements.sheetCsvUrlInput.value = state.settings.sheetCsvUrl || "";
render();
void maybeAutoSyncSheetData();

function loadState() {
  const raw =
    window.localStorage.getItem(STORAGE_KEY) ||
    window.localStorage.getItem(LEGACY_STORAGE_KEY);

  if (!raw) {
    return getDefaultState();
  }

  try {
    const parsed = JSON.parse(raw);
    return {
      ...getDefaultState(),
      ...parsed,
      settings: {
        ...getDefaultState().settings,
        ...(parsed.settings || {}),
      },
    };
  } catch {
    return getDefaultState();
  }
}

function saveState() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function getDefaultState() {
  return {
    holdings: [],
    quotes: {},
    history: [],
    fx: {
      usdToGbpRate: DEFAULT_USD_TO_GBP_RATE,
      updatedAt: "",
    },
    settings: {
      sheetCsvUrl: "",
    },
  };
}

function handleSheetSettingsSubmit(event) {
  event.preventDefault();
  const normalizedUrl = normalizeSheetCsvUrl(elements.sheetCsvUrlInput.value.trim());
  state.settings.sheetCsvUrl = normalizedUrl;
  elements.sheetCsvUrlInput.value = normalizedUrl;
  saveState();

  if (normalizedUrl) {
    setSheetStatus("Google Sheets source saved. Use Sync now to import the latest snapshot.", "success");
  } else {
    setSheetStatus("Add a published Google Sheets CSV URL so the app can import snapshot prices.", "error");
  }
}

function handleTableClick(event) {
  const deleteButton = event.target.closest("[data-delete-id]");
  const updateButton = event.target.closest("[data-update-id]");

  if (updateButton) {
    updateHoldingShares(updateButton.getAttribute("data-update-id"));
    return;
  }

  if (!deleteButton) {
    return;
  }

  const holdingId = deleteButton.getAttribute("data-delete-id");
  state.holdings = state.holdings.filter((holding) => holding.id !== holdingId);
  saveState();
  render();
}

function handleTableKeydown(event) {
  const input = event.target.closest("[data-shares-id]");

  if (!input || event.key !== "Enter") {
    return;
  }

  event.preventDefault();
  updateHoldingShares(input.getAttribute("data-shares-id"));
}

function render() {
  const rows = state.holdings.map((holding) => {
    const rawQuote = state.quotes[holding.ticker] || {
      price: 0,
      previousClose: 0,
    };
    const quote = convertQuoteToGbp(holding.ticker, rawQuote);
    const currentValue = holding.shares * quote.price;
    const previousValue = holding.shares * quote.previousClose;
    const dayChange = currentValue - previousValue;
    const dayChangePercent =
      previousValue > 0 ? (dayChange / previousValue) * 100 : 0;

    return {
      ...holding,
      quote,
      currentValue,
      dayChange,
      dayChangePercent,
    };
  });

  const totalValue = rows.reduce((sum, row) => sum + row.currentValue, 0);
  const totalChange = rows.reduce((sum, row) => sum + row.dayChange, 0);
  const previousPortfolioValue = totalValue - totalChange;
  const totalChangePercent =
    previousPortfolioValue > 0 ? (totalChange / previousPortfolioValue) * 100 : 0;
  const lastUpdatedAt = rows.reduce((latest, row) => {
    const updatedAt = row.quote.updatedAt || "";
    if (!updatedAt) {
      return latest;
    }
    return !latest || updatedAt > latest ? updatedAt : latest;
  }, "");
  const historyChanged = syncPortfolioHistory(totalValue, lastUpdatedAt);
  const chartHistory = getThirtyDayHistory();

  elements.totalValue.textContent = formatCurrency(totalValue);
  elements.totalChange.textContent = formatSignedCurrency(totalChange);
  elements.totalChange.className = getSummaryClass(totalChange);
  elements.totalChangePercent.textContent = `${formatSignedPercent(totalChangePercent)}`;
  elements.totalChangePercent.className = getSummaryClass(totalChange);
  elements.summaryDate.textContent = lastUpdatedAt
    ? `Prices last updated ${formatUpdatedTimestamp(lastUpdatedAt)}`
    : `As of ${formatSummaryDate(new Date())}`;
  renderSummaryChart(chartHistory);

  elements.tableBody.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td data-label="Company">
            <div class="company-cell">
              <strong>${escapeHtml(row.company)}</strong>
            </div>
          </td>
          <td data-label="Shares">
            <div class="shares-editor">
              <input
                class="table-shares-input"
                type="number"
                min="0.0001"
                step="0.0001"
                value="${row.shares}"
                data-shares-id="${row.id}"
                aria-label="Shares for ${escapeHtml(row.ticker)}"
              />
              <button class="icon-button table-icon-button" type="button" data-update-id="${row.id}" aria-label="Update shares for ${escapeHtml(row.ticker)}" title="Update shares">
                <span aria-hidden="true">✓</span>
              </button>
            </div>
          </td>
          <td data-label="Price">${formatCurrency(row.quote.price)}</td>
          <td data-label="Value">${formatCurrency(row.currentValue)}</td>
          <td data-label="Day change">
            <span class="change-pill ${getChangeClass(row.dayChange)}">
              ${formatSignedCurrency(row.dayChange)}
            </span>
          </td>
          <td data-label="Actions">
            <div class="row-actions">
              <button class="icon-button table-icon-button delete-icon-button" type="button" data-delete-id="${row.id}" aria-label="Remove ${escapeHtml(row.ticker)}" title="Remove holding">
                <span aria-hidden="true">×</span>
              </button>
            </div>
          </td>
        </tr>
      `
    )
    .join("");

  elements.tableFoot.innerHTML = rows.length
    ? `
        <tr class="totals-row">
          <td class="totals-label" colspan="3">Portfolio totals</td>
          <td class="totals-value" data-label="Value total">${formatCurrency(totalValue)}</td>
          <td data-label="Day change total">
            <span class="change-pill ${getChangeClass(totalChange)}">
              ${formatSignedCurrency(totalChange)}
            </span>
          </td>
          <td data-label="Actions"></td>
        </tr>
      `
    : "";

  elements.emptyState.classList.toggle("hidden", rows.length > 0);
  elements.holdingsLastUpdated.classList.toggle("hidden", !lastUpdatedAt);
  elements.holdingsLastUpdated.textContent = lastUpdatedAt
    ? `Last updated: ${formatUpdatedTimestamp(lastUpdatedAt)}`
    : "";
  elements.refreshAll.disabled = !state.settings.sheetCsvUrl && state.holdings.length === 0;
  if (historyChanged) {
    saveState();
  }
}

function clearSheetSource() {
  state.settings.sheetCsvUrl = "";
  elements.sheetCsvUrlInput.value = "";
  saveState();
  setSheetStatus("Stored Google Sheets source cleared from this browser.", "success");
}

async function refreshAllQuotes() {
  if (state.settings.sheetCsvUrl) {
    await syncSheetData({ manual: true });
    return;
  }

  setSheetStatus("Add a Google Sheets source to update prices.", "error");
}

async function syncSheetDataNow() {
  await syncSheetData({ manual: true });
}

async function maybeAutoSyncSheetData() {
  if (!state.settings.sheetCsvUrl) {
    return;
  }

  const lastUpdatedAt = state.holdings.reduce((latest, holding) => {
    const updatedAt = state.quotes[holding.ticker]?.updatedAt || "";
    return !latest || updatedAt > latest ? updatedAt : latest;
  }, "");

  if (isTimestampFromToday(lastUpdatedAt)) {
    return;
  }

  await syncSheetData({ manual: false });
}

async function syncSheetData({ manual }) {
  if (!state.settings.sheetCsvUrl) {
    setSheetStatus("Add a published Google Sheets CSV URL first.", "error");
    return;
  }

  setSheetStatus("Importing the latest Google Sheets snapshot...", "pending");
  setBusyState(true);

  try {
    const response = await fetch(state.settings.sheetCsvUrl);

    if (!response.ok) {
      throw new Error("Could not load the published Google Sheets CSV.");
    }

    const csvText = await resolveTemporaryRedirectCsv(await response.text());
    const rows = parseCsvRows(csvText);
    const imported = importSheetSnapshot(rows);

    saveState();
    render();

    setSheetStatus(
      `Updated prices for ${imported} holding${imported === 1 ? "" : "s"} from Google Sheets.`,
      "success"
    );
    if (manual) {
      highlightHoldingsSection();
    }
  } catch (error) {
    if (error instanceof TypeError && window.location.protocol === "file:") {
      setSheetStatus(
        "This browser is blocking Google Sheets fetches from a local file. Open the hosted app or run it from a local web server instead of file://.",
        "error"
      );
    } else {
      setSheetStatus(error.message, "error");
    }
  } finally {
    setBusyState(false);
  }
}

async function resolveTemporaryRedirectCsv(csvText) {
  const redirectMatch = String(csvText || "").match(/<A HREF="([^"]+output=csv[^"]*)">here<\/A>/i);

  if (!redirectMatch) {
    return csvText;
  }

  const response = await fetch(redirectMatch[1]);

  if (!response.ok) {
    throw new Error("Google Sheets redirected the CSV request, but the redirected file could not be loaded.");
  }

  return response.text();
}

function toggleSheetSyncPanel() {
  const isHidden = elements.sheetSyncContent.classList.contains("hidden");
  elements.sheetSyncContent.classList.toggle("hidden", !isHidden);
  elements.toggleSheetSync.setAttribute("aria-expanded", String(isHidden));
  elements.sheetSyncToggleLabel.textContent = isHidden ? "Hide" : "Show";
}

function highlightHoldingsSection() {
  elements.holdingsSection.classList.remove("flash-update");
  void elements.holdingsSection.offsetWidth;
  elements.holdingsSection.classList.add("flash-update");
  elements.holdingsSection.scrollIntoView({
    behavior: "smooth",
    block: "start",
  });
}

function updateHoldingShares(holdingId) {
  const input = elements.tableBody.querySelector(`[data-shares-id="${holdingId}"]`);

  if (!input) {
    return;
  }

  const shares = Number(input.value);

  if (!Number.isFinite(shares) || shares <= 0) {
    setSheetStatus("Enter a share count greater than 0 before updating.", "error");
    input.focus();
    return;
  }

  const holding = state.holdings.find((item) => item.id === holdingId);

  if (!holding) {
    return;
  }

  holding.shares = shares;
  saveState();
  render();
  highlightHoldingsSection();
  setSheetStatus(`${holding.company || holding.ticker} shares updated to ${shares}.`, "success");
}

function setSheetStatus(message, tone) {
  elements.sheetStatus.textContent = message;
  elements.sheetStatus.className = `status-text ${tone || ""}`.trim();
}

function setBusyState(isBusy) {
  elements.refreshAll.disabled = isBusy || (!state.settings.sheetCsvUrl && state.holdings.length === 0);
  elements.syncSheetNow.disabled = isBusy;
  elements.clearSheetSource.disabled = isBusy;
  elements.sheetSettingsForm.querySelector("button[type='submit']").disabled = isBusy;
  elements.tableBody.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });
}

function formatUpdatedTimestamp(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "Unknown";
  }

  return date.toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatSummaryDate(value) {
  return value.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function syncPortfolioHistory(totalValue, snapshotTimestamp) {
  const today = snapshotTimestamp ? getLocalDateKey(new Date(snapshotTimestamp)) : getLocalDateKey(new Date());
  const history = Array.isArray(state.history) ? state.history : [];
  const entry = history.find((item) => item.date === today);

  if (entry) {
    if (entry.value === totalValue) {
      return false;
    }
    entry.value = totalValue;
  } else {
    history.push({
      date: today,
      value: totalValue,
    });
  }

  state.history = history
    .filter((item) => typeof item.date === "string" && Number.isFinite(item.value))
    .sort((left, right) => left.date.localeCompare(right.date))
    .slice(-30);

  return true;
}

function getThirtyDayHistory() {
  return (Array.isArray(state.history) ? state.history : [])
    .filter((item) => typeof item.date === "string" && Number.isFinite(item.value))
    .sort((left, right) => left.date.localeCompare(right.date))
    .slice(-30);
}

function renderSummaryChart(history) {
  if (!history.length) {
    elements.summaryChart.innerHTML = "";
    elements.summaryChartRange.textContent = "";
    elements.summaryChartEmpty.classList.remove("hidden");
    return;
  }

  const series = history.map((item) => ({
    date: item.date,
    value: item.value,
  }));
  const values = series.map((item) => item.value);
  const minRaw = Math.min(...values);
  const maxRaw = Math.max(...values);
  const paddingValue = Math.max((maxRaw - minRaw) * 0.08, 1);
  const minValue = Math.max(0, minRaw - paddingValue);
  const maxValue = maxRaw + paddingValue;
  const span = maxValue - minValue || 1;
  const width = 320;
  const height = 140;
  const paddingLeft = 46;
  const paddingRight = 10;
  const paddingTop = 12;
  const paddingBottom = 24;
  const innerWidth = width - paddingLeft - paddingRight;
  const innerHeight = height - paddingTop - paddingBottom;

  const points = series
    .map((item, index) => {
      const x =
        series.length === 1
          ? width / 2
          : paddingLeft + (innerWidth * index) / (series.length - 1);
      const y = paddingTop + ((maxValue - item.value) / span) * innerHeight;
      return `${x},${y}`;
    })
    .join(" ");

  const areaPoints = `${paddingLeft},${height - paddingBottom} ${points} ${width - paddingRight},${
    height - paddingBottom
  }`;
  const firstValue = values[0];
  const lastValue = values[values.length - 1];
  const stroke = lastValue >= firstValue ? "#0f766e" : "#b42318";
  const fill = lastValue >= firstValue ? "rgba(15, 118, 110, 0.15)" : "rgba(180, 35, 24, 0.14)";
  const yTicks = [maxValue, (maxValue + minValue) / 2, minValue];
  const xTickIndexes = Array.from(new Set([0, Math.floor((series.length - 1) / 2), series.length - 1]));

  const yTickLines = yTicks
    .map((tick) => {
      const y = paddingTop + ((maxValue - tick) / span) * innerHeight;
      return `
        <line x1="${paddingLeft}" y1="${y}" x2="${width - paddingRight}" y2="${y}" class="chart-grid-line"></line>
        <text x="${paddingLeft - 6}" y="${y + 4}" text-anchor="end" class="chart-axis-label">${escapeHtml(
          formatCompactCurrency(tick)
        )}</text>
      `;
    })
    .join("");

  const xTickLabels = xTickIndexes
    .map((index) => {
      const x =
        series.length === 1
          ? width / 2
          : paddingLeft + (innerWidth * index) / (series.length - 1);
      return `<text x="${x}" y="${height - 6}" text-anchor="middle" class="chart-axis-label">${escapeHtml(
        formatChartDate(series[index].date)
      )}</text>`;
    })
    .join("");

  elements.summaryChart.innerHTML = `
    ${yTickLines}
    <line x1="${paddingLeft}" y1="${paddingTop}" x2="${paddingLeft}" y2="${height - paddingBottom}" class="chart-axis-line"></line>
    <line x1="${paddingLeft}" y1="${height - paddingBottom}" x2="${width - paddingRight}" y2="${height - paddingBottom}" class="chart-axis-line"></line>
    <polygon points="${areaPoints}" fill="${fill}"></polygon>
    <polyline points="${points}" fill="none" stroke="${stroke}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></polyline>
    ${xTickLabels}
  `;
  elements.summaryChartRange.textContent = `${formatChartDate(history[0].date)} to ${formatChartDate(
    history[history.length - 1].date
  )}`;
  elements.summaryChartEmpty.classList.add("hidden");
}

function getLocalDateKey(value) {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatChartDate(value) {
  const [year, month, day] = value.split("-");
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
  });
}

function formatCompactCurrency(value) {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function normalizeSheetCsvUrl(value) {
  if (!value) {
    return "";
  }

  try {
    const url = new URL(value);

    if (url.hostname !== "docs.google.com") {
      return value;
    }

    if (url.pathname.includes("/pubhtml")) {
      return `https://docs.google.com${url.pathname.replace("/pubhtml", "/pub")}?output=csv`;
    }

    if (url.pathname.includes("/pub")) {
      return `https://docs.google.com${url.pathname}?output=csv`;
    }

    const match = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    if (!match) {
      return value;
    }

    const gid = url.searchParams.get("gid") || "0";
    return `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv&gid=${gid}`;
  } catch {
    return value;
  }
}

function parseCsvRows(value) {
  const rows = [];
  let current = "";
  let row = [];
  let insideQuotes = false;

  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];
    const nextCharacter = value[index + 1];

    if (character === '"') {
      if (insideQuotes && nextCharacter === '"') {
        current += '"';
        index += 1;
      } else {
        insideQuotes = !insideQuotes;
      }
      continue;
    }

    if (character === "," && !insideQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !insideQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      row.push(current);
      if (row.some((cell) => cell.trim() !== "")) {
        rows.push(row);
      }
      row = [];
      current = "";
      continue;
    }

    current += character;
  }

  row.push(current);
  if (row.some((cell) => cell.trim() !== "")) {
    rows.push(row);
  }

  return rows;
}

function importSheetSnapshot(rows) {
  if (!rows.length) {
    throw new Error("The published CSV is empty.");
  }

  const headers = rows[0].map((header) => String(header || "").trim().toLowerCase());
  const requiredHeaders = ["ticker", "price"];

  for (const header of requiredHeaders) {
    if (!headers.includes(header)) {
      throw new Error(`The sheet is missing the required "${header}" column.`);
    }
  }

  const tickerIndex = headers.indexOf("ticker");
  const priceIndex = headers.indexOf("price");
  const previousCloseIndex = headers.indexOf("previousclose");
  const companyIndex = headers.indexOf("company");
  const currencyIndex = headers.indexOf("currency");
  const dateIndex = headers.indexOf("date");
  const existingHoldings = new Map(
    state.holdings.map((holding) => [holding.ticker.trim().toUpperCase(), holding])
  );
  const nextHoldings = [];
  let importedCount = 0;

  for (const cells of rows.slice(1)) {
    const ticker = String(cells[tickerIndex] || "").trim().toUpperCase();
    if (!ticker) {
      continue;
    }

    const price = Number(cells[priceIndex]);
    const rawPreviousClose = previousCloseIndex >= 0 ? Number(cells[previousCloseIndex]) : NaN;

    if (!Number.isFinite(price) || price < 0) {
      continue;
    }

    const existingHolding = existingHoldings.get(ticker);
    const existingQuote = state.quotes[ticker];
    const updatedAt = dateIndex >= 0 ? coerceSnapshotDate(cells[dateIndex]) : new Date().toISOString();
    const isSameDayUpdate = isTimestampOnSameLocalDay(existingQuote?.updatedAt, updatedAt);
    const derivedPreviousClose = Number.isFinite(rawPreviousClose)
      ? rawPreviousClose
      : isSameDayUpdate
        ? existingQuote?.previousClose
        : existingQuote?.price;
    const previousClose = Number.isFinite(derivedPreviousClose) ? derivedPreviousClose : price;
    const currency =
      currencyIndex >= 0
        ? String(cells[currencyIndex] || "").trim().toUpperCase() || getTickerCurrency(ticker)
        : existingQuote?.currency || getTickerCurrency(ticker);
    const company =
      companyIndex >= 0
        ? String(cells[companyIndex] || "").trim() || existingHolding?.company || ticker
        : existingHolding?.company || ticker;

    nextHoldings.push({
      id: existingHolding?.id || crypto.randomUUID(),
      company,
      ticker,
      shares: existingHolding?.shares || 0,
    });

    state.quotes[ticker] = {
      price,
      previousClose,
      currency,
      updatedAt,
      source: "sheet",
    };

    importedCount += 1;
  }

  if (!importedCount) {
    throw new Error("No usable holdings were found in the published CSV.");
  }

  state.holdings = nextHoldings;

  return importedCount;
}

function coerceSnapshotDate(value) {
  const text = String(value || "").trim();
  if (!text) {
    return new Date().toISOString();
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return new Date().toISOString();
  }

  return parsed.toISOString();
}

function getTickerCurrency(ticker) {
  return ticker.endsWith(".L") || ticker.endsWith(".LON") ? "GBP" : "USD";
}

function isSameLocalDay(left, right) {
  return (
    left.getFullYear() === right.getFullYear() &&
    left.getMonth() === right.getMonth() &&
    left.getDate() === right.getDate()
  );
}

function isTimestampFromToday(value) {
  if (!value) {
    return false;
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return false;
  }

  return isSameLocalDay(date, new Date());
}

function isTimestampOnSameLocalDay(left, right) {
  if (!left || !right) {
    return false;
  }

  const leftDate = new Date(left);
  const rightDate = new Date(right);

  if (Number.isNaN(leftDate.getTime()) || Number.isNaN(rightDate.getTime())) {
    return false;
  }

  return isSameLocalDay(leftDate, rightDate);
}

function getUsdToGbpRate() {
  return Number.isFinite(state.fx?.usdToGbpRate) && state.fx.usdToGbpRate > 0
    ? state.fx.usdToGbpRate
    : DEFAULT_USD_TO_GBP_RATE;
}

function convertAmountToGbp(value, currency) {
  if (!Number.isFinite(value)) {
    return 0;
  }

  if (currency === "USD") {
    return value * getUsdToGbpRate();
  }

  return value;
}

function convertQuoteToGbp(ticker, quote) {
  const currency = quote.currency || getTickerCurrency(ticker);

  return {
    ...quote,
    currency,
    price: convertAmountToGbp(quote.price, currency),
    previousClose: convertAmountToGbp(quote.previousClose, currency),
  };
}

function formatCurrency(value) {
  return currencyFormatter.format(value || 0);
}

function formatSignedCurrency(value) {
  const formatted = currencyFormatter.format(Math.abs(value || 0));
  if (value > 0) {
    return `+${formatted}`;
  }
  if (value < 0) {
    return `-${formatted}`;
  }
  return formatted;
}

function formatSignedPercent(value) {
  const formatted = `${percentFormatter.format(Math.abs(value || 0))}%`;
  if (value > 0) {
    return `+${formatted}`;
  }
  if (value < 0) {
    return `-${formatted}`;
  }
  return formatted;
}

function getChangeClass(value) {
  if (value > 0) {
    return "gain";
  }
  if (value < 0) {
    return "loss";
  }
  return "flat";
}

function getSummaryClass(value) {
  const className = getChangeClass(value);
  return `summary-value ${className}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
