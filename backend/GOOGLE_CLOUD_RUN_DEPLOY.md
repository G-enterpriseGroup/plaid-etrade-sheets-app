# Google Cloud Run Deploy - Portfolio Link Market Backend

Best Google-native route for the Python market engine.

Colab is not recommended for this app because it is a notebook runtime, not a stable API server. Cloud Run gives you a stable HTTPS URL that Apps Script can call every time.

## One-time setup

1. Go to Google Cloud Console.
2. Create or select a project.
3. Enable billing on the project.
4. Enable these APIs:
   - Cloud Run API
   - Cloud Build API
   - Artifact Registry API

## Deploy from Cloud Shell

Open Cloud Shell and run:

```bash
PROJECT_ID="YOUR_PROJECT_ID"
REGION="us-central1"
SERVICE="portfolio-link-market-backend"
TOKEN="replace-with-a-long-private-token"

gcloud config set project "$PROJECT_ID"

gcloud run deploy "$SERVICE" \
  --source ./backend \
  --region "$REGION" \
  --allow-unauthenticated \
  --set-env-vars BACKEND_API_TOKEN="$TOKEN",MARKET_LOOKBACK="2y" \
  --memory 1Gi \
  --cpu 1 \
  --timeout 300 \
  --max-instances 2
```

Cloud Run will return a service URL like:

```text
https://portfolio-link-market-backend-xxxxx-uc.a.run.app
```

## Connect Apps Script

In Apps Script > Project Settings > Script Properties, add:

```text
MARKET_BACKEND_URL=https://portfolio-link-market-backend-xxxxx-uc.a.run.app
MARKET_BACKEND_TOKEN=replace-with-a-long-private-token
```

Then run:

```text
testMarketBackendConnection
```

Then run:

```text
buildMarketAnalysisReport
```

## Why Cloud Run instead of Colab

Cloud Run is designed for stable HTTP services. Apps Script can call it through UrlFetchApp. Colab is good for manual notebooks, but it is not good as a production API because sessions disconnect, notebooks stop, and public tunnel URLs can change.
