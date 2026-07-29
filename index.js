/* Kings of Sales Trip Calculator — premium web edition */

const STORAGE_KEY = "kings-of-sales-trip-calculator-v2";
const INPUT_IDS = [
  "ClientName", "TripName", "QuoteDate", "Pax", "Currency",
  "HowManyAirports", "AirportPriceLE", "Sightseeing", "SightseeingValueLE",
  "OtherTransfersLE", "GuideLE", "LeaderLE", "LunchLE", "TicketsLE",
  "CairoNights", "CairoAccommodationUSD", "LuxorNights", "LuxorAccommodationUSD",
  "AswanNights", "AswanAccommodationUSD", "HurghadaNights", "HurghadaAccommodationUSD",
  "CruiseNights", "CruiseAccommodationUSD", "FlightsUSD", "OtherOptionsUSD",
  "InternationalFlightsUSD", "ProfitPercentage", "TotalDestination", "ProfitDestPercentage"
];

const moneyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

let saveTimer;
let toastTimer;

document.addEventListener("DOMContentLoaded", initCalculator);

function initCalculator() {
  setDefaultQuoteDate();
  restoreDraft();
  refreshQuote();

  const form = document.getElementById("tripCalculator");
  form.addEventListener("submit", (event) => event.preventDefault());

  INPUT_IDS.forEach((id) => {
    const input = document.getElementById(id);
    if (!input) return;

    input.addEventListener("input", () => {
      refreshQuote();
      scheduleDraftSave();
    });

    input.addEventListener("change", () => {
      refreshQuote();
      scheduleDraftSave();
    });
  });

  document.getElementById("calculateButton").addEventListener("click", () => {
    refreshQuote();
    showToast("Quote refreshed with the latest values.");
  });

  document.getElementById("resetButton").addEventListener("click", clearQuote);
  document.getElementById("copyButton").addEventListener("click", copyQuoteSummary);
  document.getElementById("exportButton").addEventListener("click", exportExcel);
}

function getNumber(id) {
  const input = document.getElementById(id);
  const rawValue = String(input?.value ?? "").trim().replace(",", ".");
  const value = Number.parseFloat(rawValue);
  return Number.isFinite(value) ? value : 0;
}

function getText(id) {
  return String(document.getElementById(id)?.value ?? "").trim();
}

function collectFormData() {
  return {
    clientName: getText("ClientName"),
    tripName: getText("TripName"),
    quoteDate: getText("QuoteDate"),
    pax: getNumber("Pax"),
    currency: getNumber("Currency"),
    airports: getNumber("HowManyAirports"),
    airportPrice: getNumber("AirportPriceLE"),
    sightseeing: getNumber("Sightseeing"),
    sightseeingValue: getNumber("SightseeingValueLE"),
    otherTransfers: getNumber("OtherTransfersLE"),
    guide: getNumber("GuideLE"),
    leader: getNumber("LeaderLE"),
    lunch: getNumber("LunchLE"),
    tickets: getNumber("TicketsLE"),
    cairoNights: getNumber("CairoNights"),
    cairoRate: getNumber("CairoAccommodationUSD"),
    luxorNights: getNumber("LuxorNights"),
    luxorRate: getNumber("LuxorAccommodationUSD"),
    aswanNights: getNumber("AswanNights"),
    aswanRate: getNumber("AswanAccommodationUSD"),
    redSeaNights: getNumber("HurghadaNights"),
    redSeaRate: getNumber("HurghadaAccommodationUSD"),
    cruiseNights: getNumber("CruiseNights"),
    cruiseRate: getNumber("CruiseAccommodationUSD"),
    domesticFlights: getNumber("FlightsUSD"),
    otherOptions: getNumber("OtherOptionsUSD"),
    internationalFlights: getNumber("InternationalFlightsUSD"),
    egyptProfitPercent: getNumber("ProfitPercentage"),
    destinationBase: getNumber("TotalDestination"),
    destinationProfitPercent: getNumber("ProfitDestPercentage")
  };
}

function calculateTrip() {
  const data = collectFormData();
  const safeDivide = (numerator, denominator) => denominator > 0 ? numerator / denominator : 0;

  const transfers = safeDivide(
    (data.airportPrice * data.airports) +
    (data.sightseeing * data.sightseeingValue) +
    data.otherTransfers,
    data.currency * data.pax
  );

  const gratuities = safeDivide(
    (data.guide * data.sightseeing) + data.leader,
    data.currency * data.pax
  );

  const accommodation =
    (data.cairoRate * data.cairoNights) +
    (data.luxorRate * data.luxorNights) +
    (data.aswanRate * data.aswanNights) +
    (data.redSeaRate * data.redSeaNights) +
    (data.cruiseRate * data.cruiseNights);

  // Kept intentionally identical to the model in the original calculator:
  // lunches and tickets are treated as per-traveler costs and only converted to USD.
  const localExpenses = safeDivide(data.lunch + data.tickets, data.currency);
  const egyptBase = data.domesticFlights + localExpenses + accommodation + gratuities + transfers;
  const egyptMargin = egyptBase * (data.egyptProfitPercent / 100);
  const egyptSubtotal = egyptBase + egyptMargin;
  const egyptTotal = egyptSubtotal + data.otherOptions;
  const destinationMargin = data.destinationBase * (data.destinationProfitPercent / 100);
  const destinationTotal = data.destinationBase + destinationMargin;
  const finalTotal = egyptTotal + data.internationalFlights + destinationTotal;

  return {
    data,
    isReady: data.pax > 0 && data.currency > 0,
    transfers,
    gratuities,
    accommodation,
    localExpenses,
    egyptBase,
    egyptMargin,
    egyptSubtotal,
    egyptTotal,
    destinationMargin,
    destinationTotal,
    finalTotal
  };
}

function refreshQuote() {
  const calculation = calculateTrip();
  updateQuoteUI(calculation);
  return calculation;
}

function updateQuoteUI(calculation) {
  const { data } = calculation;
  const clientLine = data.clientName
    ? `Prepared for ${data.clientName}${data.tripName ? ` · ${data.tripName}` : ""}`
    : data.tripName || "Add your client details to begin.";

  document.getElementById("summaryClient").textContent = clientLine;
  document.getElementById("result").textContent = calculation.isReady ? formatMoney(calculation.finalTotal) : "—";
  document.getElementById("egyptTotal").textContent = formatMoney(calculation.egyptTotal);
  document.getElementById("internationalTotal").textContent = formatMoney(data.internationalFlights);
  document.getElementById("destinationTotal").textContent = formatMoney(calculation.destinationTotal);
  document.getElementById("accommodationTotal").textContent = formatMoney(calculation.accommodation);
  document.getElementById("egyptBaseTotal").textContent = formatMoney(calculation.egyptBase);
  document.getElementById("egyptMarginTotal").textContent = formatMoney(calculation.egyptMargin);

  const priceHint = document.getElementById("priceHint");
  const status = document.getElementById("statusMessage");

  if (!calculation.isReady) {
    priceHint.textContent = "Enter travelers and the EGP / USD rate.";
    status.textContent = "Set the two required fields to activate the live price.";
    return;
  }

  priceHint.textContent = `Based on ${formatTravelerCount(data.pax)} and your current cost inputs.`;
  status.textContent = "All totals are live. Your draft saves automatically in this browser.";
}

function formatMoney(value) {
  return moneyFormatter.format(Number.isFinite(value) ? value : 0);
}

function formatTravelerCount(value) {
  const formatted = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(value);
  return `${formatted} ${value === 1 ? "traveler" : "travelers"}`;
}

function setDefaultQuoteDate() {
  const dateField = document.getElementById("QuoteDate");
  if (dateField && !dateField.value) {
    dateField.value = new Date().toISOString().slice(0, 10);
  }
}

function scheduleDraftSave() {
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveDraft, 180);
}

function saveDraft() {
  const draft = Object.fromEntries(INPUT_IDS.map((id) => [id, document.getElementById(id)?.value ?? ""]));

  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(draft));
  } catch (error) {
    // Private browsing or a restrictive browser can block storage. The calculator still works normally.
    console.warn("Could not save quote draft", error);
  }
}

function restoreDraft() {
  try {
    const savedDraft = window.localStorage.getItem(STORAGE_KEY);
    if (!savedDraft) return;

    const draft = JSON.parse(savedDraft);
    INPUT_IDS.forEach((id) => {
      if (typeof draft[id] === "string" && document.getElementById(id)) {
        document.getElementById(id).value = draft[id];
      }
    });
  } catch (error) {
    console.warn("Could not restore quote draft", error);
  }
}

function clearQuote() {
  document.getElementById("tripCalculator").reset();
  setDefaultQuoteDate();

  try {
    window.localStorage.removeItem(STORAGE_KEY);
  } catch (error) {
    console.warn("Could not clear saved quote draft", error);
  }

  refreshQuote();
  showToast("Quote cleared. You can start a fresh calculation.");
}

async function copyQuoteSummary() {
  const calculation = refreshQuote();
  if (!calculation.isReady) {
    showToast("Add travelers and the EGP / USD rate before copying the quote.");
    return;
  }

  const { data } = calculation;
  const quoteLines = [
    "KINGS OF SALES — TRIP QUOTE",
    data.clientName ? `Client: ${data.clientName}` : null,
    data.tripName ? `Trip: ${data.tripName}` : null,
    `Travelers: ${formatTravelerCount(data.pax)}`,
    "",
    `Egypt total: ${formatMoney(calculation.egyptTotal)}`,
    `International flights: ${formatMoney(data.internationalFlights)}`,
    `Destination add-on: ${formatMoney(calculation.destinationTotal)}`,
    "",
    `PRICE PER TRAVELER: ${formatMoney(calculation.finalTotal)}`
  ].filter(Boolean).join("\n");

  try {
    await navigator.clipboard.writeText(quoteLines);
    showToast("Quote summary copied to your clipboard.");
  } catch (error) {
    const helper = document.createElement("textarea");
    helper.value = quoteLines;
    helper.setAttribute("readonly", "");
    helper.style.position = "fixed";
    helper.style.opacity = "0";
    document.body.appendChild(helper);
    helper.select();
    document.execCommand("copy");
    helper.remove();
    showToast("Quote summary copied to your clipboard.");
  }
}

async function exportExcel() {
  const calculation = refreshQuote();
  if (!calculation.isReady) {
    showToast("Add travelers and the EGP / USD rate before exporting Excel.");
    return;
  }

  if (!window.ExcelJS || !window.saveAs) {
    showToast("Excel export is still loading. Please try again in a moment.");
    return;
  }

  const exportButton = document.getElementById("exportButton");
  const originalText = exportButton.innerHTML;
  exportButton.disabled = true;
  exportButton.innerHTML = "<span aria-hidden=\"true\">…</span> Preparing Excel";

  try {
    const workbook = buildWorkbook(calculation);
    const buffer = await workbook.xlsx.writeBuffer();
    const dateLabel = calculation.data.quoteDate || new Date().toISOString().slice(0, 10);
    const clientLabel = makeFileSafe(calculation.data.clientName || calculation.data.tripName || "TripQuote");
    const filename = `KingsOfSales_${clientLabel}_${dateLabel}.xlsx`;

    saveAs(
      new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      filename
    );
    showToast("Smart Excel exported — its formulas will recalculate when values change.");
  } catch (error) {
    console.error("Excel export failed", error);
    showToast("Excel could not be exported. Please try again.");
  } finally {
    exportButton.disabled = false;
    exportButton.innerHTML = originalText;
  }
}

function buildWorkbook(calculation) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Shreef Ammar";
  workbook.lastModifiedBy = "Kings of Sales Trip Calculator";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties = {
    calcMode: "auto",
    fullCalcOnLoad: true,
    forceFullCalc: true
  };

  const inputSheet = workbook.addWorksheet("Quote Inputs", {
    views: [{ state: "frozen", ySplit: 4, showGridLines: false }]
  });
  const inputRows = buildInputSheet(inputSheet, calculation.data);

  const calculationSheet = workbook.addWorksheet("Calculation", {
    views: [{ state: "frozen", ySplit: 4, showGridLines: false }]
  });
  const calculationRows = buildCalculationSheet(calculationSheet, inputRows, calculation);

  const summarySheet = workbook.addWorksheet("Quote Summary", {
    views: [{ showGridLines: false }]
  });
  buildSummarySheet(summarySheet, inputRows, calculationRows, calculation);

  workbook.worksheets.forEach((sheet) => {
    sheet.pageSetup = {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
      margins: { left: 0.28, right: 0.28, top: 0.38, bottom: 0.38, header: 0.1, footer: 0.1 }
    };
  });

  return workbook;
}

function buildInputSheet(sheet, data) {
  const colors = excelColors();
  const rows = {};

  sheet.columns = [
    { width: 20 },
    { width: 34 },
    { width: 22 },
    { width: 20 }
  ];

  sheet.mergeCells("A1:D1");
  const titleCell = sheet.getCell("A1");
  titleCell.value = "KINGS OF SALES  |  QUOTE INPUTS";
  styleExcelTitle(titleCell, colors);
  sheet.getRow(1).height = 32;

  sheet.mergeCells("A2:D2");
  const guideCell = sheet.getCell("A2");
  guideCell.value = "Edit only the blue cells in this sheet. The Calculation and Quote Summary sheets update automatically in Excel.";
  guideCell.font = { name: "Aptos", size: 10, italic: true, color: { argb: colors.muted } };
  guideCell.fill = fill(colors.softBlue);
  guideCell.alignment = { horizontal: "left", vertical: "middle" };
  sheet.getRow(2).height = 25;

  const headerRow = 4;
  ["SECTION", "INPUT", "EDITABLE VALUE", "UNIT / NOTE"].forEach((title, index) => {
    const cell = sheet.getCell(headerRow, index + 1);
    cell.value = title;
    cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.white } };
    cell.fill = fill(colors.teal);
    cell.alignment = { horizontal: index === 2 ? "center" : "left", vertical: "middle" };
    cell.border = border(colors.teal);
  });
  sheet.getRow(headerRow).height = 23;

  let row = 5;
  const addSection = (title) => {
    sheet.mergeCells(`A${row}:D${row}`);
    const cell = sheet.getCell(`A${row}`);
    cell.value = title;
    cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.navy } };
    cell.fill = fill(colors.goldLight);
    cell.alignment = { vertical: "middle" };
    cell.border = border(colors.goldLight);
    sheet.getRow(row).height = 22;
    row += 1;
  };

  const addInput = (key, category, label, value, unit, type = "number") => {
    rows[key] = row;
    const cells = [
      sheet.getCell(`A${row}`), sheet.getCell(`B${row}`), sheet.getCell(`C${row}`), sheet.getCell(`D${row}`)
    ];
    cells.forEach((cell) => {
      cell.font = { name: "Aptos", size: 10, color: { argb: colors.navy } };
      cell.fill = fill(colors.white);
      cell.border = border(colors.grid);
      cell.alignment = { vertical: "middle", horizontal: "left" };
    });

    cells[0].value = category;
    cells[0].font = { name: "Aptos", size: 9, bold: true, color: { argb: colors.muted } };
    cells[1].value = label;
    cells[2].value = type === "date" ? excelDate(value) : value;
    cells[2].fill = fill(colors.inputBlue);
    cells[2].font = { name: "Aptos", size: 10, bold: type !== "text", color: { argb: colors.navy } };
    cells[2].alignment = { vertical: "middle", horizontal: type === "text" || type === "date" ? "left" : "right" };
    cells[3].value = unit;
    cells[3].font = { name: "Aptos", size: 9, color: { argb: colors.muted } };

    if (type === "currency" || type === "number") cells[2].numFmt = "#,##0.00";
    if (type === "whole") cells[2].numFmt = "0";
    if (type === "percent") cells[2].numFmt = "0.00\"%\"";
    if (type === "date") cells[2].numFmt = "dd mmm yyyy";

    if (["currency", "number", "whole", "percent"].includes(type)) {
      cells[2].dataValidation = numericValidation(type === "whole" ? 1 : 0, type === "whole");
    }

    sheet.getRow(row).height = 21;
    row += 1;
  };

  addSection("QUOTE DETAILS");
  addInput("clientName", "Quote", "Client name", data.clientName, "Text", "text");
  addInput("tripName", "Quote", "Trip name", data.tripName, "Text", "text");
  addInput("quoteDate", "Quote", "Quote date", data.quoteDate, "Date", "date");
  addInput("pax", "Quote", "Travelers", data.pax, "Guests", "whole");
  addInput("currency", "Quote", "EGP / USD exchange rate", data.currency, "EGP per USD", "number");

  row += 1;
  addSection("EGYPT OPERATIONS  |  LOCAL EGP COSTS");
  addInput("airportPrice", "Transfers", "Airport transfer value", data.airportPrice, "EGP", "currency");
  addInput("airports", "Transfers", "Airport transfers", data.airports, "Movements", "whole");
  addInput("sightseeing", "Touring", "Day tours", data.sightseeing, "Tours", "whole");
  addInput("sightseeingValue", "Touring", "Day-tour transfer value", data.sightseeingValue, "EGP", "currency");
  addInput("otherTransfers", "Transfers", "Other transfers", data.otherTransfers, "EGP", "currency");
  addInput("guide", "Operations", "Guide cost", data.guide, "EGP", "currency");
  addInput("leader", "Operations", "Tour leader cost", data.leader, "EGP", "currency");
  addInput("lunch", "Expenses", "Lunches", data.lunch, "EGP", "currency");
  addInput("tickets", "Expenses", "Tickets", data.tickets, "EGP", "currency");

  row += 1;
  addSection("ACCOMMODATION  |  USD PER TRAVELER");
  addInput("cairoNights", "Cairo", "Cairo nights", data.cairoNights, "Nights", "whole");
  addInput("cairoRate", "Cairo", "Cairo rate per night", data.cairoRate, "USD", "currency");
  addInput("luxorNights", "Luxor", "Luxor nights", data.luxorNights, "Nights", "whole");
  addInput("luxorRate", "Luxor", "Luxor rate per night", data.luxorRate, "USD", "currency");
  addInput("aswanNights", "Aswan", "Aswan nights", data.aswanNights, "Nights", "whole");
  addInput("aswanRate", "Aswan", "Aswan rate per night", data.aswanRate, "USD", "currency");
  addInput("redSeaNights", "Red Sea", "Red Sea nights", data.redSeaNights, "Nights", "whole");
  addInput("redSeaRate", "Red Sea", "Red Sea rate per night", data.redSeaRate, "USD", "currency");
  addInput("cruiseNights", "Nile Cruise", "Nile Cruise nights", data.cruiseNights, "Nights", "whole");
  addInput("cruiseRate", "Nile Cruise", "Nile Cruise rate per night", data.cruiseRate, "USD", "currency");

  row += 1;
  addSection("AIR & COMMERCIAL  |  USD PER TRAVELER");
  addInput("domesticFlights", "Flights", "Domestic flights", data.domesticFlights, "USD", "currency");
  addInput("otherOptions", "Options", "Other options", data.otherOptions, "USD", "currency");
  addInput("internationalFlights", "Flights", "International flights", data.internationalFlights, "USD", "currency");
  addInput("egyptProfitPercent", "Margin", "Egypt profit", data.egyptProfitPercent, "%", "percent");

  row += 1;
  addSection("ADD-ON DESTINATION  |  USD PER TRAVELER");
  addInput("destinationBase", "Destination", "Destination base cost", data.destinationBase, "USD", "currency");
  addInput("destinationProfitPercent", "Margin", "Destination profit", data.destinationProfitPercent, "%", "percent");

  return rows;
}

function buildCalculationSheet(sheet, inputRows, calculation) {
  const colors = excelColors();
  const rows = {};
  const inputRef = (key) => `'Quote Inputs'!$C$${inputRows[key]}`;

  sheet.columns = [
    { width: 19 },
    { width: 34 },
    { width: 20 },
    { width: 46 }
  ];

  sheet.mergeCells("A1:D1");
  const titleCell = sheet.getCell("A1");
  titleCell.value = "KINGS OF SALES  |  LIVE CALCULATION";
  styleExcelTitle(titleCell, colors);
  sheet.getRow(1).height = 32;

  sheet.mergeCells("A2:D2");
  const guideCell = sheet.getCell("A2");
  guideCell.value = "Every value in column C is a live Excel formula. Change blue input cells on Quote Inputs and this sheet recalculates.";
  guideCell.font = { name: "Aptos", size: 10, italic: true, color: { argb: colors.muted } };
  guideCell.fill = fill(colors.softBlue);
  guideCell.alignment = { horizontal: "left", vertical: "middle" };
  sheet.getRow(2).height = 25;

  ["AREA", "PRICE COMPONENT", "USD / TRAVELER", "CALCULATION NOTE"].forEach((title, index) => {
    const cell = sheet.getCell(4, index + 1);
    cell.value = title;
    cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.white } };
    cell.fill = fill(colors.teal);
    cell.alignment = { horizontal: index === 2 ? "center" : "left", vertical: "middle" };
    cell.border = border(colors.teal);
  });
  sheet.getRow(4).height = 23;

  let row = 5;
  const addFormula = (key, area, label, formula, result, note, emphasis = false) => {
    rows[key] = row;
    const rowCells = [1, 2, 3, 4].map((column) => sheet.getCell(row, column));
    rowCells.forEach((cell) => {
      cell.border = border(colors.grid);
      cell.fill = fill(emphasis ? colors.summaryLight : colors.white);
      cell.font = { name: "Aptos", size: 10, color: { argb: colors.navy } };
      cell.alignment = { vertical: "middle", horizontal: "left" };
    });

    rowCells[0].value = area;
    rowCells[0].font = { name: "Aptos", size: 9, bold: true, color: { argb: colors.muted } };
    rowCells[1].value = label;
    rowCells[1].font = { name: "Aptos", size: 10, bold: emphasis, color: { argb: colors.navy } };
    rowCells[2].value = { formula, result: roundForExcel(result) };
    rowCells[2].numFmt = "$#,##0.00;[Red]-$#,##0.00";
    rowCells[2].fill = fill(emphasis ? colors.goldLight : colors.formulaBlue);
    rowCells[2].font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.navy } };
    rowCells[2].alignment = { vertical: "middle", horizontal: "right" };
    rowCells[3].value = note;
    rowCells[3].font = { name: "Aptos", size: 9, color: { argb: colors.muted } };
    sheet.getRow(row).height = emphasis ? 24 : 21;
    row += 1;
  };

  addFormula("domesticFlights", "Egypt", "Domestic flights", inputRef("domesticFlights"), calculation.data.domesticFlights, "Linked from Quote Inputs");
  addFormula(
    "transfers",
    "Egypt",
    "Transfers per traveler",
    `IFERROR((${inputRef("airportPrice")}*${inputRef("airports")}+${inputRef("sightseeing")}*${inputRef("sightseeingValue")}+${inputRef("otherTransfers")})/${inputRef("currency")}/${inputRef("pax")},0)`,
    calculation.transfers,
    "Airport + day-tour + other transfers, converted and shared by travelers"
  );
  addFormula(
    "gratuities",
    "Egypt",
    "Guide & leader per traveler",
    `IFERROR((${inputRef("guide")}*${inputRef("sightseeing")}+${inputRef("leader")})/${inputRef("currency")}/${inputRef("pax")},0)`,
    calculation.gratuities,
    "Guide/day tours + leader, converted and shared by travelers"
  );
  addFormula(
    "localExpenses",
    "Egypt",
    "Lunches & tickets",
    `IFERROR((${inputRef("lunch")}+${inputRef("tickets")})/${inputRef("currency")},0)`,
    calculation.localExpenses,
    "Local EGP costs converted to USD"
  );
  addFormula(
    "accommodation",
    "Egypt",
    "Accommodation",
    `${inputRef("cairoNights")}*${inputRef("cairoRate")}+${inputRef("luxorNights")}*${inputRef("luxorRate")}+${inputRef("aswanNights")}*${inputRef("aswanRate")}+${inputRef("redSeaNights")}*${inputRef("redSeaRate")}+${inputRef("cruiseNights")}*${inputRef("cruiseRate")}`,
    calculation.accommodation,
    "Nights × rate for Cairo, Luxor, Aswan, Red Sea and Nile Cruise"
  );
  addFormula(
    "egyptBase",
    "Egypt",
    "Egypt operating base",
    `SUM(C${rows.domesticFlights}:C${rows.accommodation})`,
    calculation.egyptBase,
    "Sum of flights, transfers, guide, local expenses and stays",
    true
  );
  addFormula(
    "egyptMargin",
    "Egypt",
    "Egypt profit",
    `C${rows.egyptBase}*${inputRef("egyptProfitPercent")}/100`,
    calculation.egyptMargin,
    "Operating base × Egypt profit percentage"
  );
  addFormula(
    "egyptSubtotal",
    "Egypt",
    "Egypt subtotal before options",
    `C${rows.egyptBase}+C${rows.egyptMargin}`,
    calculation.egyptSubtotal,
    "Egypt base plus profit"
  );
  addFormula("otherOptions", "Egypt", "Other options", inputRef("otherOptions"), calculation.data.otherOptions, "Added after Egypt profit, per original model");
  addFormula(
    "egyptTotal",
    "Egypt",
    "Egypt total",
    `C${rows.egyptSubtotal}+C${rows.otherOptions}`,
    calculation.egyptTotal,
    "Final Egypt price per traveler",
    true
  );

  row += 1;
  addFormula("internationalFlights", "Flights", "International flights", inputRef("internationalFlights"), calculation.data.internationalFlights, "Added with no profit margin");

  row += 1;
  addFormula("destinationBase", "Destination", "Destination base cost", inputRef("destinationBase"), calculation.data.destinationBase, "Linked from Quote Inputs");
  addFormula(
    "destinationMargin",
    "Destination",
    "Destination profit",
    `C${rows.destinationBase}*${inputRef("destinationProfitPercent")}/100`,
    calculation.destinationMargin,
    "Destination base × destination profit percentage"
  );
  addFormula(
    "destinationTotal",
    "Destination",
    "Destination total",
    `C${rows.destinationBase}+C${rows.destinationMargin}`,
    calculation.destinationTotal,
    "Destination base plus profit",
    true
  );

  row += 1;
  addFormula(
    "grandTotal",
    "Final Quote",
    "FINAL TOTAL PER TRAVELER",
    `C${rows.egyptTotal}+C${rows.internationalFlights}+C${rows.destinationTotal}`,
    calculation.finalTotal,
    "Egypt total + international flights + destination total",
    true
  );

  return rows;
}

function buildSummarySheet(sheet, inputRows, calculationRows, calculation) {
  const colors = excelColors();
  const inputRef = (key) => `'Quote Inputs'!$C$${inputRows[key]}`;
  const calculationRef = (key) => `Calculation!$C$${calculationRows[key]}`;

  sheet.columns = [
    { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }
  ];

  sheet.mergeCells("A1:F1");
  const title = sheet.getCell("A1");
  title.value = "KINGS OF SALES";
  title.font = { name: "Aptos Display", size: 20, bold: true, color: { argb: colors.white } };
  title.fill = fill(colors.navy);
  title.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 36;

  sheet.mergeCells("A2:F2");
  const subtitle = sheet.getCell("A2");
  subtitle.value = "PREMIUM TRIP QUOTATION  |  LIVE EXCEL MODEL";
  subtitle.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.gold } };
  subtitle.fill = fill(colors.navy);
  subtitle.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 22;

  addSummaryDetail(sheet, "A4", "B4:F4", "CLIENT", inputRef("clientName"), calculation.data.clientName, colors);
  addSummaryDetail(sheet, "A5", "B5:F5", "TRIP", inputRef("tripName"), calculation.data.tripName, colors);
  addSummaryDetail(sheet, "A6", "B6:F6", "QUOTE DATE", inputRef("quoteDate"), excelDate(calculation.data.quoteDate), colors, "dd mmm yyyy");
  addSummaryDetail(sheet, "A7", "B7:F7", "TRAVELERS", inputRef("pax"), calculation.data.pax, colors, "0");

  sheet.mergeCells("A9:F9");
  const totalLabel = sheet.getCell("A9");
  totalLabel.value = "ESTIMATED PRICE PER TRAVELER";
  totalLabel.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.navy } };
  totalLabel.fill = fill(colors.goldLight);
  totalLabel.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(9).height = 24;

  sheet.mergeCells("A10:F13");
  const totalValue = sheet.getCell("A10");
  totalValue.value = { formula: calculationRef("grandTotal"), result: roundForExcel(calculation.finalTotal) };
  totalValue.numFmt = "$#,##0.00;[Red]-$#,##0.00";
  totalValue.font = { name: "Aptos Display", size: 30, bold: true, color: { argb: colors.navy } };
  totalValue.fill = fill(colors.goldPale);
  totalValue.alignment = { horizontal: "center", vertical: "middle" };
  [10, 11, 12, 13].forEach((row) => { sheet.getRow(row).height = 22; });

  const tableHeaders = ["PRICE BREAKDOWN", "USD / TRAVELER"];
  sheet.mergeCells("A15:D15");
  sheet.mergeCells("E15:F15");
  const leftHeader = sheet.getCell("A15");
  leftHeader.value = tableHeaders[0];
  const rightHeader = sheet.getCell("E15");
  rightHeader.value = tableHeaders[1];
  [leftHeader, rightHeader].forEach((cell) => {
    cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.white } };
    cell.fill = fill(colors.teal);
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  sheet.getRow(15).height = 23;

  addSummaryAmount(sheet, 16, "Egypt total", calculationRef("egyptTotal"), calculation.egyptTotal, colors);
  addSummaryAmount(sheet, 17, "International flights", calculationRef("internationalFlights"), calculation.data.internationalFlights, colors);
  addSummaryAmount(sheet, 18, "Destination add-on", calculationRef("destinationTotal"), calculation.destinationTotal, colors);

  sheet.mergeCells("A21:F21");
  const note = sheet.getCell("A21");
  note.value = "To update this quote in Excel: open Quote Inputs, edit the blue cells, then review the refreshed totals here.";
  note.font = { name: "Aptos", size: 10, italic: true, color: { argb: colors.muted } };
  note.fill = fill(colors.softBlue);
  note.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getRow(21).height = 34;

  sheet.mergeCells("A23:F23");
  const footer = sheet.getCell("A23");
  footer.value = "Designed by Shreef Ammar";
  footer.font = { name: "Aptos", size: 9, italic: true, color: { argb: colors.muted } };
  footer.alignment = { horizontal: "center", vertical: "middle" };
}

function addSummaryDetail(sheet, labelAddress, mergedValueRange, label, formula, result, colors, numFmt) {
  const labelCell = sheet.getCell(labelAddress);
  labelCell.value = label;
  labelCell.font = { name: "Aptos", size: 9, bold: true, color: { argb: colors.muted } };
  labelCell.fill = fill(colors.white);
  labelCell.alignment = { vertical: "middle" };
  labelCell.border = border(colors.grid);

  sheet.mergeCells(mergedValueRange);
  const valueCell = sheet.getCell(mergedValueRange.split(":")[0]);
  valueCell.value = { formula, result };
  valueCell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.navy } };
  valueCell.fill = fill(colors.white);
  valueCell.alignment = { vertical: "middle", horizontal: "left" };
  valueCell.border = border(colors.grid);
  if (numFmt) valueCell.numFmt = numFmt;

  const row = Number(labelAddress.replace(/[^0-9]/g, ""));
  sheet.getRow(row).height = 21;
}

function addSummaryAmount(sheet, row, label, formula, result, colors) {
  sheet.mergeCells(`A${row}:D${row}`);
  sheet.mergeCells(`E${row}:F${row}`);
  const labelCell = sheet.getCell(`A${row}`);
  const valueCell = sheet.getCell(`E${row}`);

  labelCell.value = label;
  labelCell.font = { name: "Aptos", size: 10, color: { argb: colors.navy } };
  labelCell.fill = fill(row % 2 === 0 ? colors.white : colors.rowTint);
  labelCell.alignment = { horizontal: "left", vertical: "middle" };
  labelCell.border = border(colors.grid);

  valueCell.value = { formula, result: roundForExcel(result) };
  valueCell.numFmt = "$#,##0.00;[Red]-$#,##0.00";
  valueCell.font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.navy } };
  valueCell.fill = fill(row % 2 === 0 ? colors.white : colors.rowTint);
  valueCell.alignment = { horizontal: "right", vertical: "middle" };
  valueCell.border = border(colors.grid);
  sheet.getRow(row).height = 23;
}

function excelColors() {
  return {
    navy: "FF102B46",
    teal: "FF176D72",
    gold: "FFD49A43",
    goldLight: "FFF6E4BB",
    goldPale: "FFFFF6DF",
    white: "FFFFFFFF",
    softBlue: "FFF0F6FB",
    inputBlue: "FFDCECF9",
    formulaBlue: "FFE8F1F8",
    summaryLight: "FFF7FBFD",
    rowTint: "FFF7FAFC",
    muted: "FF627487",
    grid: "FFD6E0E8"
  };
}

function styleExcelTitle(cell, colors) {
  cell.font = { name: "Aptos Display", size: 16, bold: true, color: { argb: colors.white } };
  cell.fill = fill(colors.navy);
  cell.alignment = { horizontal: "left", vertical: "middle" };
}

function fill(color) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: color } };
}

function border(color) {
  return {
    top: { style: "thin", color: { argb: color } },
    left: { style: "thin", color: { argb: color } },
    bottom: { style: "thin", color: { argb: color } },
    right: { style: "thin", color: { argb: color } }
  };
}

function numericValidation(minimum, wholeNumber = false) {
  return {
    type: wholeNumber ? "whole" : "decimal",
    operator: "greaterThanOrEqual",
    allowBlank: true,
    formulae: [minimum],
    showErrorMessage: true,
    errorStyle: "stop",
    errorTitle: "Invalid value",
    error: `Enter a number greater than or equal to ${minimum}.`
  };
}

function excelDate(dateString) {
  if (!dateString) return null;
  const [year, month, day] = dateString.split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day, 12);
}

function roundForExcel(value) {
  return Number(Number.isFinite(value) ? value.toFixed(8) : 0);
}

function makeFileSafe(value) {
  const clean = value
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 48);
  return clean || "TripQuote";
}

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("is-visible");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.remove("is-visible"), 3200);
}
