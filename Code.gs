function fetchStockPrices() {
  const apiKey = "your_api_key"; // Replace with Alpaca API key
  const secretKey = "your_secret_key"; // Replace with Alpaca secret key 
  const symbols = ["CRWD", "PLTR", "SOFI"];
  const barsUrl = "https://data.alpaca.markets/v2/stocks/bars";
  const quotesUrl = "https://data.alpaca.markets/v2/stocks/quotes/latest";
  const tradesUrl = "https://data.alpaca.markets/v2/stocks/trades/latest";

  const headers = {
    "APCA-API-KEY-ID": apiKey,
    "APCA-API-SECRET-KEY": secretKey
  };

  // Safe 30-day window for SIP
  const endDate = new Date(); // now UTC
  const startDate = new Date(endDate.getTime() - 30 * 24 * 60 * 60 * 1000);
  const startISO = startDate.toISOString();
  const endISO = endDate.toISOString();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  const header = [
    "Symbol",
    "Latest NBBO Mid / Trade (IEX)",
    "Quote/Trade Time (CT)",
    "SIP Close (latest 1D)",
    "SIP Time (CT)"
  ];
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");

  let row = 2;

  symbols.forEach((symbol) => {
    let liveMidOrTrade = "Error", liveTs = "Error";
    let sipClose = "need to upgrade for this", sipTs = "need to upgrade for this";

    // ---- IEX latest quote NBBO mid; fallback to latest trade ----
    try {
      const qRes = UrlFetchApp.fetch(`${quotesUrl}?${encodeParams({ symbols: symbol, feed: "iex" })}`, {
        headers, muteHttpExceptions: true
      });
      if (qRes.getResponseCode() !== 200) throw new Error(`HTTP ${qRes.getResponseCode()}: ${qRes.getContentText()}`);
      const qData = JSON.parse(qRes.getContentText());
      const q = qData.quotes && qData.quotes[symbol];

      if (q && q.bp != null && q.ap != null && q.bp > 0 && q.ap > 0) {
        liveMidOrTrade = (q.bp + q.ap) / 2;
        liveTs = formatCT(new Date(q.t));
      } else {
        const tRes = UrlFetchApp.fetch(`${tradesUrl}?${encodeParams({ symbols: symbol, feed: "iex" })}`, {
          headers, muteHttpExceptions: true
        });
        if (tRes.getResponseCode() !== 200) throw new Error(`HTTP ${tRes.getResponseCode()}: ${tRes.getContentText()}`);
        const tData = JSON.parse(tRes.getContentText());
        const tr = tData.trades && tData.trades[symbol];
        if (!tr || tr.p == null) throw new Error("No usable quote or trade (IEX)");
        liveMidOrTrade = tr.p;
        liveTs = formatCT(new Date(tr.t));
      }
    } catch (e) {
      Logger.log(`Quotes/trades (IEX) error for ${symbol}: ${e}`);
      liveMidOrTrade = "Error";
      liveTs = errMsg(e);
    }

    // ---- SIP daily bar (latest) ----
    try {
      const paramsBarsSIP = {
        symbols: symbol,
        timeframe: "1Day",
        start: startISO,
        end: endISO,
        limit: 1,
        feed: "sip"
      };
      const resS = UrlFetchApp.fetch(`${barsUrl}?${encodeParams(paramsBarsSIP)}`, {
        headers, muteHttpExceptions: true
      });
      if (resS.getResponseCode() !== 200) throw new Error(`HTTP ${resS.getResponseCode()}: ${resS.getContentText()}`);

      const dataS = JSON.parse(resS.getContentText());
      const barsS = (dataS && dataS.bars && dataS.bars[symbol]) || [];
      if (barsS.length === 0) throw new Error("No bar data (SIP)");

      const lastSip = barsS[0];
      sipClose = lastSip.c;
      sipTs = formatCT(new Date(lastSip.t));
    } catch (e) {
      Logger.log(`SIP bars error for ${symbol}: ${e}`);
      sipClose = "need to upgrade for this";
      sipTs = "need to upgrade for this";
    }

    // Write row
    sheet.getRange(row, 1, 1, header.length).setValues([
      [symbol, liveMidOrTrade, liveTs, sipClose, sipTs]
    ]);

    formatMoney(sheet.getRange(row, 2));
    formatMoney(sheet.getRange(row, 4));

    row++;
  });

  sheet.autoResizeColumns(1, header.length);
}

// ---------- Helpers ----------
function encodeParams(params) {
  return Object.keys(params)
    .map((key) => `${key}=${encodeURIComponent(params[key])}`)
    .join("&");
}

function formatMoney(range) {
  range.setNumberFormat("$0.00").setHorizontalAlignment("center");
}

function formatCT(dateObj) {
  return Utilities.formatDate(dateObj, "America/Chicago", "MM/dd/yyyy hh:mm:ss a").toLowerCase();
}

function errMsg(e) {
  return (e && e.message) ? e.message : String(e);
}
