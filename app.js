const STORAGE_KEY = "share-tracker-state-v2";
const LEGACY_STORAGE_KEY = "share-tracker-state-v1";
const ALPHA_VANTAGE_URL = "https://www.alphavantage.co/query";
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
  settingsForm: document.getElementById("settings-form"),
  apiKeyInput: document.getElementById("api-key-input"),
  apiStatus: document.getElementById("api-status"),
  holdingForm: document.getElementById("holding-form"),
  holdingStatus: document.getElementById("holding-status"),
  clearHoldingForm: document.getElementById("clear-holding-form"),
  toggleMarketData: document.getElementById("toggle-market-data"),
  marketDataContent: document.getElementById("market-data-content"),
  marketDataToggleLabel: document.getElementById("market-data-toggle-label"),
  toggleManualQuote: document.getElementById("toggle-manual-quote"),
  manualQuoteContent: document.getElementById("manual-quote-content"),
  manualQuoteToggleLabel: document.getElementById("manual-quote-toggle-label"),
  quoteForm: document.getElementById("quote-form"),
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
  clearApiKey: document.getElementById("clear-api-key"),
};

const state = loadState();

elements.settingsForm.addEventListener("submit", handleSettingsSubmit);
elements.holdingForm.addEventListener("submit", handleHoldingSubmit);
elements.quoteForm.addEventListener("submit", handleQuoteSubmit);
elements.tableBody.addEventListener("click", handleTableClick);
elements.tableBody.addEventListener("keydown", handleTableKeydown);
elements.refreshAll.addEventListener("click", refreshAllQuotes);
elements.clearApiKey.addEventListener("click", clearApiKey);
elements.clearHoldingForm.addEventListener("click", resetHoldingForm);
elements.toggleMarketData.addEventListener("click", toggleMarketDataPanel);
elements.toggleManualQuote.addEventListener("click", toggleManualQuotePanel);

elements.apiKeyInput.value = state.settings.apiKey;
render();
void maybeAutoRefreshStaleData();

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
    historicalQuotes: {},
    historicalUpdatedAt: {},
    fx: {
      usdToGbpRate: DEFAULT_USD_TO_GBP_RATE,
      updatedAt: "",
    },
    settings: {
      apiKey: "",
    },
  };
}

function handleSettingsSubmit(event) {
  event.preventDefault();
  state.settings.apiKey = elements.apiKeyInput.value.trim();
  saveState();

  if (state.settings.apiKey) {
    setApiStatus("API key saved in this browser.", "success");
  } else {
    setApiStatus("Add an API key so the app can fetch quotes.", "error");
  }
}

function handleHoldingSubmit(event) {
  event.preventDefault();

  const formData = new FormData(event.currentTarget);
  const ticker = String(formData.get("ticker") || "").trim().toUpperCase();
  const shares = Number(formData.get("shares"));

  if (!isTickerFormatValid(ticker)) {
    setHoldingStatus("Enter a valid ticker format, for example AAPL or VOD.L.", "error");
    return;
  }

  if (!Number.isFinite(shares) || shares <= 0) {
    setHoldingStatus("Enter a share count greater than 0.", "error");
    return;
  }

  const existingHolding = state.holdings.find((holding) => holding.ticker === ticker);

  if (existingHolding) {
    existingHolding.shares += shares;
  } else {
    state.holdings.push({
      id: crypto.randomUUID(),
      company: ticker,
      ticker,
      shares,
    });
  }

  event.currentTarget.reset();
  saveState();
  render();
  highlightHoldingsSection();

  if (state.settings.apiKey) {
    setHoldingStatus(
      `${ticker} saved. Refreshing portfolio breakdown data...`,
      "success"
    );
    refreshNewHoldingData(ticker);
  } else {
    setHoldingStatus(
      `${ticker} saved. Add an API key if you want live quote lookups too.`,
      "success"
    );
  }
}

async function refreshNewHoldingData(ticker) {
  try {
    await refreshExchangeRates([ticker]);
    await fetchAndStoreQuote(ticker, { force: false });
    await fetchAndStoreHistoricalQuotes(ticker, { force: false });
    saveState();
    render();
    highlightHoldingsSection();
    setHoldingStatus(`${ticker} saved and portfolio breakdown refreshed.`, "success");
    setApiStatus(`${ticker} updated from Alpha Vantage, with USD holdings converted to GBP.`, "success");
  } catch (error) {
    setHoldingStatus(
      `${ticker} saved, but live quote refresh did not complete. You can refresh holdings later.`,
      "error"
    );
    setApiStatus(error.message, "error");
  }
}

function handleQuoteSubmit(event) {
  event.preventDefault();

  const formData = new FormData(event.currentTarget);
  const ticker = String(formData.get("ticker") || "").trim().toUpperCase();
  const price = Number(formData.get("price"));
  const previousClose = Number(formData.get("previousClose"));

  if (
    !ticker ||
    !Number.isFinite(price) ||
    !Number.isFinite(previousClose) ||
    price < 0 ||
    previousClose < 0
  ) {
    return;
  }

  state.quotes[ticker] = {
    price,
    previousClose,
    currency: getTickerCurrency(ticker),
    updatedAt: new Date().toISOString(),
    source: "manual",
  };

  event.currentTarget.reset();
  saveState();
  render();
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
  const historyChanged = syncPortfolioHistory(totalValue);
  const chartHistory = getChartHistory(rows);

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
              ${row.quote.source === "manual" ? '<span class="holding-meta-manual">Manual quote</span>' : ""}
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
  elements.refreshAll.disabled = state.holdings.length === 0;
  if (historyChanged) {
    saveState();
  }
}

function clearApiKey() {
  state.settings.apiKey = "";
  elements.apiKeyInput.value = "";
  saveState();
  setApiStatus("Stored API key cleared from this browser.", "success");
  setHoldingStatus("Ticker validation will now use format checks only.", "pending");
}

async function refreshAllQuotes() {
  if (!state.settings.apiKey) {
    setApiStatus("Add your Alpha Vantage API key first.", "error");
    return;
  }

  if (state.holdings.length === 0) {
    setApiStatus("Add at least one holding before refreshing quotes.", "error");
    return;
  }

  setBusyState(true);
  setApiStatus(`Refreshing ${state.holdings.length} holding${state.holdings.length === 1 ? "" : "s"}...`, "pending");

  let refreshed = 0;
  let cached = 0;
  let historyRefreshed = 0;

  try {
    await refreshExchangeRates();

    for (const holding of state.holdings) {
      const didRefresh = await fetchAndStoreQuote(holding.ticker, { force: false });
      if (didRefresh) {
        refreshed += 1;
      } else {
        cached += 1;
      }
    }

    historyRefreshed = await refreshHistoricalSeries();

    saveState();
    render();
    const parts = [];
    parts.push(
      refreshed > 0
        ? `Updated ${refreshed} holding${refreshed === 1 ? "" : "s"}`
        : "No holdings needed a quote refresh"
    );
    if (cached > 0) {
      parts.push(`${cached} used cached prices`);
    }
    if (historyRefreshed > 0) {
      parts.push(`${historyRefreshed} historical series refreshed`);
    }
    setApiStatus(`${parts.join(". ")}.`, "success");
  } catch (error) {
    setApiStatus(error.message, "error");
  } finally {
    setBusyState(false);
  }
}

async function maybeAutoRefreshStaleData() {
  if (!state.settings.apiKey || state.holdings.length === 0) {
    return;
  }

  const staleTickers = state.holdings
    .map((holding) => holding.ticker)
    .filter((ticker) => !isQuoteFresh(ticker) || !isHistoricalSeriesFresh(ticker));
  const needsFxRefresh = state.holdings.some((holding) => getTickerCurrency(holding.ticker) === "USD") && !isFxRateFresh();

  if (!staleTickers.length && !needsFxRefresh) {
    return;
  }

  setBusyState(true);
  setApiStatus("Checking for fresh market data...", "pending");

  let refreshedQuotes = 0;
  let refreshedHistory = 0;

  try {
    await refreshExchangeRates(
      needsFxRefresh ? state.holdings.map((holding) => holding.ticker) : staleTickers
    );

    for (const ticker of staleTickers) {
      const didRefreshQuote = await fetchAndStoreQuote(ticker, { force: false });
      const didRefreshHistory = await fetchAndStoreHistoricalQuotes(ticker, { force: false });

      if (didRefreshQuote) {
        refreshedQuotes += 1;
      }

      if (didRefreshHistory) {
        refreshedHistory += 1;
      }
    }

    saveState();
    render();

    const messages = [];
    if (refreshedQuotes > 0) {
      messages.push(`Updated ${refreshedQuotes} holding${refreshedQuotes === 1 ? "" : "s"}`);
    }
    if (refreshedHistory > 0) {
      messages.push(`refreshed ${refreshedHistory} chart histor${refreshedHistory === 1 ? "y" : "ies"}`);
    }

    setApiStatus(
      messages.length ? `${messages.join(" and ")} automatically.` : "Market data was already up to date.",
      "success"
    );
  } catch (error) {
    setApiStatus(error.message, "error");
  } finally {
    setBusyState(false);
  }
}

async function refreshQuoteForTicker(ticker) {
  if (!state.settings.apiKey) {
    setApiStatus("Add your Alpha Vantage API key first.", "error");
    return;
  }

  setBusyState(true);
  setApiStatus(`Refreshing ${ticker}...`, "pending");

  try {
    await refreshExchangeRates([ticker]);
    const didRefreshQuote = await fetchAndStoreQuote(ticker, { force: false });
    const didRefreshHistory = await fetchAndStoreHistoricalQuotes(ticker, { force: false });
    saveState();
    render();
    if (didRefreshQuote || didRefreshHistory) {
      setApiStatus(`${ticker} updated from Alpha Vantage.`, "success");
    } else {
      setApiStatus(`${ticker} is already up to date, so the cached data was reused.`, "success");
    }
  } catch (error) {
    setApiStatus(error.message, "error");
  } finally {
    setBusyState(false);
  }
}

async function fetchAndStoreQuote(ticker, options = {}) {
  if (!options.force && isQuoteFresh(ticker)) {
    return false;
  }

  const apiTicker = normalizeTickerForApi(ticker);
  const params = new URLSearchParams({
    function: "GLOBAL_QUOTE",
    symbol: apiTicker,
    apikey: state.settings.apiKey,
  });
  const response = await fetch(`${ALPHA_VANTAGE_URL}?${params.toString()}`);

  if (!response.ok) {
    throw new Error(`Alpha Vantage request failed for ${ticker}.`);
  }

  const payload = await response.json();

  if (payload.Note) {
    throw new Error("Alpha Vantage rate limit reached. Please wait a minute and try again.");
  }

  if (payload.Information) {
    throw new Error(payload.Information);
  }

  if (payload["Error Message"]) {
    throw new Error(`Ticker ${ticker} was not recognised by Alpha Vantage.`);
  }

  const quote = payload["Global Quote"];
  const price = Number(quote?.["05. price"]);
  const previousClose = Number(quote?.["08. previous close"]);
  const currency = getTickerCurrency(ticker);

  if (!quote || !Number.isFinite(price) || !Number.isFinite(previousClose)) {
    throw new Error(`No usable quote came back for ${ticker}. Check the ticker or try again later.`);
  }

  state.quotes[ticker] = {
    price,
    previousClose,
    currency,
    updatedAt: new Date().toISOString(),
    source: "alpha-vantage",
  };

  return true;
}

async function fetchAndStoreHistoricalQuotes(ticker, options = {}) {
  if (!options.force && isHistoricalSeriesFresh(ticker)) {
    return false;
  }

  const apiTicker = normalizeTickerForApi(ticker);
  const params = new URLSearchParams({
    function: "TIME_SERIES_DAILY_ADJUSTED",
    symbol: apiTicker,
    outputsize: "compact",
    apikey: state.settings.apiKey,
  });
  const response = await fetch(`${ALPHA_VANTAGE_URL}?${params.toString()}`);

  if (!response.ok) {
    throw new Error(`Historical data request failed for ${ticker}.`);
  }

  const payload = await response.json();

  if (payload.Note) {
    throw new Error("Alpha Vantage rate limit reached while loading historical data. Please wait a minute and try again.");
  }

  if (payload.Information) {
    throw new Error(payload.Information);
  }

  if (payload["Error Message"]) {
    throw new Error(`Historical data for ${ticker} was not recognised by Alpha Vantage.`);
  }

  const series =
    payload["Time Series (Daily)"] ||
    payload["Time Series (Daily) Adjusted"] ||
    payload["Time Series (Daily) "];

  if (!series || typeof series !== "object") {
    throw new Error(`No usable historical data came back for ${ticker}.`);
  }

  state.historicalQuotes[ticker] = Object.entries(series)
    .map(([date, values]) => ({
      date,
      close: Number(values["5. adjusted close"] || values["4. close"]),
    }))
    .filter((entry) => Number.isFinite(entry.close))
    .sort((left, right) => left.date.localeCompare(right.date))
    .slice(-30);
  state.historicalUpdatedAt[ticker] = new Date().toISOString();

  return true;
}

async function refreshHistoricalSeries() {
  let refreshed = 0;

  for (const holding of state.holdings) {
    const didRefresh = await fetchAndStoreHistoricalQuotes(holding.ticker, { force: false });
    if (didRefresh) {
      refreshed += 1;
    }
  }

  return refreshed;
}

async function refreshExchangeRates(tickers = state.holdings.map((holding) => holding.ticker)) {
  const hasUsdHoldings = tickers.some((ticker) => getTickerCurrency(ticker) === "USD");

  if (!hasUsdHoldings || !state.settings.apiKey || isFxRateFresh()) {
    return;
  }

  const params = new URLSearchParams({
    function: "CURRENCY_EXCHANGE_RATE",
    from_currency: "USD",
    to_currency: "GBP",
    apikey: state.settings.apiKey,
  });
  const response = await fetch(`${ALPHA_VANTAGE_URL}?${params.toString()}`);

  if (!response.ok) {
    throw new Error("Could not refresh the USD to GBP exchange rate.");
  }

  const payload = await response.json();

  if (payload.Note) {
    throw new Error("Alpha Vantage rate limit reached while loading USD to GBP exchange rates. Please wait a minute and try again.");
  }

  if (payload.Information) {
    throw new Error(payload.Information);
  }

  const rate = Number(payload["Realtime Currency Exchange Rate"]?.["5. Exchange Rate"]);

  if (!Number.isFinite(rate) || rate <= 0) {
    throw new Error("No usable USD to GBP exchange rate came back from Alpha Vantage.");
  }

  state.fx.usdToGbpRate = rate;
  state.fx.updatedAt = new Date().toISOString();
}

async function verifyTicker(ticker) {
  try {
    const apiTicker = normalizeTickerForApi(ticker);
    const params = new URLSearchParams({
      function: "SYMBOL_SEARCH",
      keywords: apiTicker,
      apikey: state.settings.apiKey,
    });
    const response = await fetch(`${ALPHA_VANTAGE_URL}?${params.toString()}`);

    if (!response.ok) {
      return {
        valid: false,
        message: `Could not verify ${ticker} right now.`,
      };
    }

    const payload = await response.json();

    if (payload.Note) {
      return {
        valid: false,
        message: "Alpha Vantage rate limit reached while verifying the ticker. Please wait a minute and try again.",
      };
    }

    if (payload.Information || payload["Error Message"]) {
      return {
        valid: false,
        message: `Alpha Vantage could not verify ${ticker}.`,
      };
    }

    const match = Array.isArray(payload.bestMatches)
      ? payload.bestMatches.find((item) => {
          const symbol = String(item["1. symbol"] || "").toUpperCase();
          return symbol === ticker || symbol === apiTicker;
        })
      : null;

    if (!match) {
      return {
        valid: false,
        message: `Ticker ${ticker} was not found in Alpha Vantage symbol search.`,
      };
    }

    return {
      valid: true,
      name: String(match["2. name"] || "").trim(),
    };
  } catch {
    return {
      valid: false,
      message: `Could not verify ${ticker}. Check your connection or try the manual quote form.`,
    };
  }
}

function resetHoldingForm() {
  elements.holdingForm.reset();
  setHoldingStatus("Holding form cleared.", "success");
}

function toggleManualQuotePanel() {
  const isHidden = elements.manualQuoteContent.classList.contains("hidden");
  elements.manualQuoteContent.classList.toggle("hidden", !isHidden);
  elements.toggleManualQuote.setAttribute("aria-expanded", String(isHidden));
  elements.manualQuoteToggleLabel.textContent = isHidden ? "Hide" : "Show";
}

function toggleMarketDataPanel() {
  const isHidden = elements.marketDataContent.classList.contains("hidden");
  elements.marketDataContent.classList.toggle("hidden", !isHidden);
  elements.toggleMarketData.setAttribute("aria-expanded", String(isHidden));
  elements.marketDataToggleLabel.textContent = isHidden ? "Hide" : "Show";
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
    setHoldingStatus("Enter a share count greater than 0 before updating.", "error");
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
  setHoldingStatus(`${holding.company || holding.ticker} shares updated to ${shares}.`, "success");
}

function setApiStatus(message, tone) {
  elements.apiStatus.textContent = message;
  elements.apiStatus.className = `status-text ${tone || ""}`.trim();
}

function setHoldingStatus(message, tone) {
  elements.holdingStatus.textContent = message;
  elements.holdingStatus.className = `status-text ${tone || ""}`.trim();
}

function setBusyState(isBusy) {
  elements.refreshAll.disabled = isBusy || state.holdings.length === 0;
  elements.clearApiKey.disabled = isBusy;
  elements.settingsForm.querySelector("button[type='submit']").disabled = isBusy;
  elements.tableBody.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });
}

function setFormBusyState(form, isBusy) {
  form.querySelectorAll("input, button").forEach((control) => {
    control.disabled = isBusy;
  });
}

function getUpdatedLabel(value) {
  if (!value) {
    return "No quote yet";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "Quote updated";
  }

  return `Updated ${date.toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  })}`;
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

function syncPortfolioHistory(totalValue) {
  const today = getLocalDateKey(new Date());
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

function getChartHistory(rows) {
  const historicalRows = rows
    .map((row) => ({
      ticker: row.ticker,
      shares: row.shares,
      series: Array.isArray(state.historicalQuotes[row.ticker]) ? state.historicalQuotes[row.ticker] : [],
    }))
    .filter((row) => row.series.length);

  if (!historicalRows.length) {
    return getThirtyDayHistory();
  }

  const allDates = Array.from(
    new Set(historicalRows.flatMap((row) => row.series.map((entry) => entry.date)))
  ).sort((left, right) => left.localeCompare(right));

  return allDates
    .map((date) => {
      let total = 0;
      let found = false;

      for (const row of historicalRows) {
        const entry = row.series.find((item) => item.date === date);
        if (entry) {
          total += convertAmountToGbp(entry.close, getTickerCurrency(row.ticker)) * row.shares;
          found = true;
        }
      }

      return found ? { date, value: total } : null;
    })
    .filter(Boolean)
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

function getSourceLabel(source) {
  if (source === "alpha-vantage") {
    return "API";
  }
  if (source === "manual") {
    return "Manual";
  }
  return "Pending";
}

function getSourceClass(source) {
  if (source === "alpha-vantage") {
    return "api";
  }
  if (source === "manual") {
    return "manual";
  }
  return "empty";
}

function isTickerFormatValid(ticker) {
  return /^[A-Z0-9.-]{1,10}$/.test(ticker);
}

function normalizeTickerForApi(ticker) {
  if (ticker.endsWith(".L")) {
    return `${ticker.slice(0, -2)}.LON`;
  }

  return ticker;
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

function isQuoteFresh(ticker) {
  const quote = state.quotes[ticker];
  if (!quote || quote.source === "manual") {
    return false;
  }

  return isTimestampFromToday(quote.updatedAt);
}

function isHistoricalSeriesFresh(ticker) {
  const series = state.historicalQuotes[ticker];
  if (!Array.isArray(series) || series.length === 0) {
    return false;
  }

  return isTimestampFromToday(state.historicalUpdatedAt?.[ticker]);
}

function isFxRateFresh() {
  return isTimestampFromToday(state.fx?.updatedAt);
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
