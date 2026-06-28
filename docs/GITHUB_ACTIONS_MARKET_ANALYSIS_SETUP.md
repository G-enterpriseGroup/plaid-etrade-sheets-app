# GitHub-Only Market Analysis Setup

This avoids Render, Railway, Fly.io, and Google Cloud Run.

## Flow

```text
Google Sheet Holdings tab
  -> Apps Script MarketAnalysisReport.gs
  -> GitHub runtime/market-inputs/request_*.json
  -> GitHub Actions workflow
  -> Python backend/market_analysis_engine.py
  -> GitHub runtime/market-outputs/report_*.json
  -> Apps Script writes Report Market Analysis tab
```

## Files

```text
.github/workflows/market-analysis.yml
backend/market_analysis_engine.py
backend/github_action_market_runner.py
backend/requirements.txt
apps-script/MarketAnalysisReport.gs
```

## Apps Script Properties

Set these in Apps Script -> Project Settings -> Script properties:

```text
GITHUB_MARKET_ACCESS = your GitHub fine-grained access value
GITHUB_REPO_FULL_NAME = G-enterpriseGroup/plaid-etrade-sheets-app
GITHUB_BRANCH = main
GITHUB_MARKET_WORKFLOW = market-analysis.yml
```

## GitHub access value permissions

Use a fine-grained GitHub access value scoped only to this repo.

Repository permissions needed:

```text
Contents: Read and write
Actions: Read and write
Metadata: Read-only
```

## Test

In Apps Script, run:

```text
testMarketGitHubConnection
```

Then run:

```text
buildMarketAnalysisReport
```

The sidebar button also calls `buildMarketAnalysisReport()`.

## Notes

GitHub Actions is not instant like a live API. The sidebar spinner may stay open 1-4 minutes while Python installs dependencies, downloads market data, commits the report JSON, and Apps Script polls for the result.

This is a repo-only development architecture. For a paid production app, a real backend is still faster and cleaner, but this removes the separate deploy step for now.
