async function loadWorkbookData() {
  const response = await fetch('./data/workbook.json?ts=' + Date.now());
  if (!response.ok) throw new Error('Failed to load workbook data');
  return await response.json();
}

loadWorkbookData().then(data => {
  console.log(data);

  const plannedDemand = data["Planned Demand"];
  const financialForecast = data["Financial Forecast"];
  const programDeepDive = data["Program Deep Dive"];
  const pmDash = data["PM Dash"];

  // build charts from these arrays
});
