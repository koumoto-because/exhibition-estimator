# Verification Guide

## Build
npm ci
npm run build

## UI Smoke Test
npm start
- App must launch without SIGABRT
- Confirm main window appears

## Functional Checks
1) Import v1.1.2 sample JSON (see scripts/sample-v112.json)
2) Switch category tabs and confirm items render
3) Open siteCosts tab, edit amounts/notes
4) Confirm siteCosts.* has {amount, currency, notes} and currency defaults to JPY
