# DSP Sheet Maker

Express microservice that renders structured content briefs into Google Sheets. Built to mirror the `markdown-doc-renderer` architecture (per-user OAuth, Postgres-backed refresh tokens, Railway deployment).

Matches the styling of the **"roulette online"** tab in `NewExample_betonline.ag - Content Briefs _ Internal.xlsx`.

## Endpoints

### `GET /`
Health check.

### `GET /connect/google`
Starts the Google OAuth flow. After approval, your refresh token is stored in `user_tokens` keyed by your email. Use that email as `userKey` in subsequent requests.

Scopes requested:
- `spreadsheets`
- `drive.file`
- `userinfo.email`, `userinfo.profile`

### `GET /oauth2/callback/google`
OAuth callback (used automatically by Google).

### `POST /render-sheet`
Renders a content brief into a target sheet.

**Body** — single job OR array of jobs:
```json
{
  "userKey": "you@example.com",
  "spreadsheetId": "16YsIxmjfm1T191-K5V1ad-YXapljYdYKiEbBKfDqce0",
  "sheetName": "Sheet1",
  "sheetId": "1275091366",
  "meta": {
    "homeUrl": "betonline.ag",
    "mainKeyword": "roulette online",
    "keywordVolume": "",
    "recommendedUrl": "https://www.betonline.ag/casino/roulette",
    "minWordCount": "4823",
    "existingOrNew": "Existing",
    "pageType": "Game page",
    "targetGeo": "US",
    "clearScopeLink": "https://www.clearscope.io/..."
  },
  "tables": {
    "contentOutline": [
      {
        "Section": "H1",
        "Heading": "Play Roulette Online for Real Money at BetOnline",
        "Writer Instructions": "...",
        "Type": "Fixed",
        "Capsule?": "N/A",
        "Word Target": "0 w",
        "Required Elements": "...",
        "Entities / Terms": "BetOnline",
        "WRITER ✓": "FALSE",
        "EDITOR ✓": "FALSE"
      }
    ],
    "seoTerms": [
      { "primary": "real money", "secondary": "real-money", "min": 2, "max": 6, "current": 0, "adj": "+2" }
    ],
    "top10Rankings": [
      { "rank": 1, "pageTitle": "...", "url": "https://...", "headerOutline": "• ..." }
    ],
    "questions": ["Q1", "Q2"],
    "trustElements": "...",
    "benefitsCta": "...",
    "painPoints": "(optional)",
    "clearscopeWriterGrade": "A++",
    "clearscopeEditorGrade": "",
    "clearscopeWriterReadability": "College",
    "clearscopeEditorReadability": ""
  },
  "notes": {
    "writerNotes": "...",
    "notesForUploader": "",
    "otherFeatures": ""
  }
}
```

You can also post `{ userKey, jobs: [job1, job2] }` to run multiple at once.

If `sheetId` (gid) is provided and matches an existing tab, it's used directly. Otherwise the service looks up by `sheetName`, and creates the tab if missing.

## Deploying to Railway

1. Create a Postgres plugin in Railway and copy its `DATABASE_URL` into the service.
2. Set env vars:
   - `DATABASE_URL` — from Railway Postgres
   - `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET` — from Google Cloud Console OAuth credentials (web application type)
   - `APP_BASE_URL` — your Railway URL, e.g. `https://dsp-sheet-maker.up.railway.app`
   - `BODY_SIZE_LIMIT` — optional, default `10mb`
3. In Google Cloud Console, add `${APP_BASE_URL}/oauth2/callback/google` as an authorized redirect URI.
4. Deploy. Then visit `${APP_BASE_URL}/connect/google` once to link your Google account.

## One-time setup per user

Each user who wants to render sheets must:
1. Visit `${APP_BASE_URL}/connect/google`
2. Approve the scopes
3. Copy the email shown on the "Connected!" page — that's their `userKey`

## Notes on the renderer

- The Google Sheets API cannot place floating images (unlike Apps Script's `insertImage`). The DSP logo is rendered via `=IMAGE(url, 1)` inside merged H1:K2, which fits-to-cell preserving aspect ratio.
- Content Outline row heights are fixed (not auto-sized) because the Sheets API has no `autoResizeDimensions` per-row call that plays well with batchUpdate + merges. If you need a different body row height, adjust in `sheet-renderer.js` → `renderContentOutline`.
- The `adj` column of SEO Terms uses a live formula `=if(F<D,...)` matching the xlsx source. The `adj` values in the payload are informational only (they're not written to the sheet).
