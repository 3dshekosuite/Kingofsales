// index.js

function getNumber(id) {
  const el = document.getElementById(id);
  return parseFloat(el?.value || "0");
}

function calculatePrice() {
  const pax = getNumber("Pax");
  const cur = getNumber("Currency");

  const transfers = (
    getNumber("AirportPriceLE") * getNumber("HowManyAirports") +
    getNumber("Sightseeing") * getNumber("SightseeingValueLE") +
    getNumber("OtherTransfersLE")
  ) / cur / pax;

  const gratuities = (
    getNumber("GuideLE") * getNumber("Sightseeing") +
    getNumber("LeaderLE")
  ) / cur / pax;

  const accommodation =
    getNumber("CairoAccommodationUSD") * getNumber("CairoNights") +
    getNumber("LuxorAccommodationUSD") * getNumber("LuxorNights") +
    getNumber("AswanAccommodationUSD") * getNumber("AswanNights") +
    getNumber("HurghadaAccommodationUSD") * getNumber("HurghadaNights") +
    getNumber("CruiseAccommodationUSD") * getNumber("CruiseNights");

  const expenses = (
    getNumber("LunchLE") + getNumber("TicketsLE")
  ) / cur;

  let egypt = getNumber("FlightsUSD") + expenses + accommodation + gratuities + transfers;
  egypt *= 1 + getNumber("ProfitPercentage") / 100;
  egypt += getNumber("OtherOptionsUSD");
  const grandBeforeIntl = egypt + getNumber("InternationalFlightsUSD");

  const destTotal = getNumber("TotalDestination");
  const destWithProfit = destTotal * (1 + getNumber("ProfitDestPercentage") / 100);

  const finalTotal = grandBeforeIntl + destWithProfit;

  document.getElementById("result").textContent =
    `Price Per Pax: ${finalTotal.toFixed(2)} USD`;
}

async function exportExcel() {
  // Préparation des données
  const rows = [
    ['Pax', getNumber('Pax')],
    ['Client Name', document.getElementById('ClientName').value || ''],
    ['Airport Price LE', getNumber('AirportPriceLE')],
    ['How Many Airports', getNumber('HowManyAirports')],
    ['Sightseeing', getNumber('Sightseeing')],
    ['Sightseeing Value LE', getNumber('SightseeingValueLE')],
    ['Other Transfers LE', getNumber('OtherTransfersLE')],
    ['Cairo Accommodation USD', getNumber('CairoAccommodationUSD')],
    ['Cairo Nights', getNumber('CairoNights')],
    ['Luxor Accommodation USD', getNumber('LuxorAccommodationUSD')],
    ['Luxor Nights', getNumber('LuxorNights')],
    ['Aswan Accommodation USD', getNumber('AswanAccommodationUSD')],
    ['Aswan Nights', getNumber('AswanNights')],
    ['Hurghada Accommodation USD', getNumber('HurghadaAccommodationUSD')],
    ['Hurghada Nights', getNumber('HurghadaNights')],
    ['Cruise Accommodation USD', getNumber('CruiseAccommodationUSD')],
    ['Cruise Nights', getNumber('CruiseNights')],
    ['Lunch LE', getNumber('LunchLE')],
    ['Tickets LE', getNumber('TicketsLE')],
    ['Flights USD', getNumber('FlightsUSD')],
    ['Other Options USD', getNumber('OtherOptionsUSD')],
    ['International Flights USD', getNumber('InternationalFlightsUSD')],
    ['Guide LE', getNumber('GuideLE')],
    ['Leader LE', getNumber('LeaderLE')],
    ['Currency', getNumber('Currency')],
    ['Profit Percentage', getNumber('ProfitPercentage')],
    ['Total Destination USD', getNumber('TotalDestination')],
    ['Profit Dest %', getNumber('ProfitDestPercentage')],
  ];

  // Recalcul pour sécurité
  const pax = getNumber("Pax"), cur = getNumber("Currency");
  const transfers = (
    getNumber("AirportPriceLE") * getNumber("HowManyAirports") +
    getNumber("Sightseeing") * getNumber("SightseeingValueLE") +
    getNumber("OtherTransfersLE")
  ) / cur / pax;
  const gratuities = (
    getNumber("GuideLE") * getNumber("Sightseeing") +
    getNumber("LeaderLE")
  ) / cur / pax;
  const accommodation =
    getNumber("CairoAccommodationUSD") * getNumber("CairoNights") +
    getNumber("LuxorAccommodationUSD") * getNumber("LuxorNights") +
    getNumber("AswanAccommodationUSD") * getNumber("AswanNights") +
    getNumber("HurghadaAccommodationUSD") * getNumber("HurghadaNights") +
    getNumber("CruiseAccommodationUSD") * getNumber("CruiseNights");
  const expenses = (getNumber("LunchLE") + getNumber("TicketsLE")) / cur;
  let egypt = getNumber("FlightsUSD") + expenses + accommodation + gratuities + transfers;
  egypt *= 1 + getNumber("ProfitPercentage")/100;
  egypt += getNumber("OtherOptionsUSD");
  const grandBeforeIntl = egypt + getNumber("InternationalFlightsUSD");
  const destWithProfit = getNumber("TotalDestination") * (1 + getNumber("ProfitDestPercentage")/100);
  const finalTotal = grandBeforeIntl + destWithProfit;

  // Création du classeur
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Trip Calculation', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  // Colonnes
  sheet.columns = [
    { header: 'Field', key: 'field', width: 30 },
    { header: 'Value', key: 'value', width: 20 },
  ];

  // Style de l'en-tête
  const headerRow = sheet.getRow(1);
  headerRow.font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2E75B6' }  // bleu moderne
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  // Ajout des données
  rows.forEach((r, i) => {
    const row = sheet.addRow({ field: r[0], value: r[1] });
    // Zébrage
    if ((i + 2) % 2 === 0) {
      row.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' } // gris clair tirant sur le bleu
      };
    }
  });

  // Ligne vide + Destination with Profit + Final Total
  sheet.addRow([]);
  const destRow = sheet.addRow({ field: 'Destination with Profit', value: destWithProfit.toFixed(2) });
  const finalRow = sheet.addRow({ field: 'Final Total Per Pax', value: finalTotal.toFixed(2) });

  // Applique bordures et police Times New Roman 12 pt
  sheet.eachRow((row, rowNumber) => {
    row.eachCell(cell => {
      cell.font = { name: 'Times New Roman', size: 12, bold: rowNumber === 1 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  // Téléchargement
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer], { type: 'application/octet-stream' }), 'TripCalculation.xlsx');
}
