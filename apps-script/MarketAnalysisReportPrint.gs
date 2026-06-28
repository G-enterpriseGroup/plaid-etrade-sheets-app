/**
 * Market Analysis Report print entrypoint.
 * Kept as a small alias so the sidebar can call one stable function name.
 * The main report engine already writes the final portrait print layout.
 */
function buildMarketAnalysisReportPrint() {
  return buildMarketAnalysisReport();
}
