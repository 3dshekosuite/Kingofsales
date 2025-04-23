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

  const grand = egypt + getNumber("InternationalFlightsUSD");

  const result = document.getElementById("result");
  result.textContent = `Price Per Pax: ${grand.toFixed(2)} USD`;
}
