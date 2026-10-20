# Pokémon Masters EX Gym Challenge Tracker

A tool to help players manage their teams and coordinate battles in the **Gym Challenge** game mode of Pokémon Masters EX.

## Overview

The Gym Challenge Tracker simplifies team management by reducing the amount of information players need to enter manually.

The application can analyze screenshots of sync pairs, match recognized names against a centralized catalog, and synchronize roster information with Google Sheets.

It is designed to work with an existing Google Sheets-based tracking system, without requiring a separate database or server.

## Features

- **Screenshot scanning:** Extract sync pair information from screenshots.
- **Sync pair recognition:** Match recognized names against the project's sync pair catalog.
- **Centralized catalog:** Load sync pair data from a JSON file hosted on GitHub.
- **Roster management:** Keep track of each player's available sync pairs.
- **Google Sheets integration:** Synchronize roster information with the existing spreadsheet.
- **Manual verification:** Review and correct uncertain matches before updating the roster.

## How It Works

1. **Capture:** The user provides a screenshot containing sync pairs.
2. **Recognition:** The application processes the screenshot and identifies potential sync pairs.
3. **Matching:** Recognized names are matched against the canonical names in the sync pair catalog.
4. **Verification:** The user reviews the results and corrects any uncertain matches.
5. **Synchronization:** The selected information is sent to Google Sheets, where the roster data is updated.

## Setup and Installation

### 1. Clone the Repository

```bash
git clone <repository-url>
cd <repository-folder>
```

### 2. Configure Google Sheets

The application requires a Google Apps Script to communicate with Google Sheets.

1. Open the Google Sheets spreadsheet you want to use.
2. Open **Extensions → Apps Script**.
3. Copy the contents of `escaner.gs` from this repository into the Apps Script editor.
4. Save the script.
5. Deploy or update the Apps Script as a web app, configuring access as required for your setup.
6. Copy the deployed web app URL and configure the application to use it.

### 3. Configure the Shared Token

A shared token is required for communication between the application and Google Apps Script.

1. Open `escaner.gs` in the Apps Script editor.
2. Find the token configuration and set your own secret token.
3. Open `app.js` in the project.
4. Set the same token in the corresponding configuration.
5. Save both files and make sure the values match exactly.

**Important:** The token must be identical in `escaner.gs` and `app.js` for synchronization to work.

Use your own token rather than a publicly shared example. Do not commit real credentials or secrets to a public repository. If `app.js` is served to users in a browser, its token can be inspected by those users; it should not be treated as a secure secret against untrusted clients.

### 4. Run the Application

Install any required dependencies and start the application using the commands specified by the project.

Make sure the Google Apps Script URL and shared token are configured correctly before attempting to synchronize data.

## Data Sources and Integration

### Sync Pair Catalog

The application loads sync pair information from a JSON catalog hosted on GitHub.

This allows the catalog to be updated independently of the application.

### Google Sheets

Google Sheets stores player roster information. The application communicates with a Google Apps Script endpoint to synchronize data.

The spreadsheet, Apps Script deployment, and token configuration must all be set up correctly for synchronization to work.

## Limitations

- Screenshot recognition may not identify every sync pair correctly.
- Similar names, alternate outfits, and different sync pair variants may require manual verification.
- Synchronization depends on the availability and correct configuration of the Google Apps Script endpoint.
- Roster accuracy depends on the correctness of the recognized and confirmed information.

## Future Improvements

Potential improvements include:

- Improving screenshot recognition accuracy.
- Enhancing synchronization error handling.
- Making roster management more convenient.
- Adding further automation to support Gym Challenge team planning.

## Disclaimer

This is an unofficial fan-made project created to assist players of Pokémon Masters EX.

It is not affiliated with, endorsed by, or associated with DeNA, The Pokémon Company, or Nintendo.

Pokémon and all related names and assets belong to their respective owners.
