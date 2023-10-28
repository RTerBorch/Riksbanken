# Currency Exchange Rate Downloader and Converter

## Introduction
This Java program downloads foreign exchange rates against the Swedish Krona (SEK) for specific dates and saves them to Excel files. The application leverages the Riksbank API for fetching the exchange rates. It supports multiple currencies, including AUD, BRL, CAD, CHF, etc.

## Features
- Downloads historical forex data from Riksbank API.
- Saves data into Excel files.
- Converts Excel data to JSON format.
- Merges the fetched data with existing datasets.

## API Endpoints

### `/getObservation` (POST)
Fetches exchange rate data for specified currencies within a date range.

#### Parameters
- **seriesIdList**: A list of currency codes (e.g., `AUD`, `BRL`, `CAD`, etc.). The seriesId will be in the format `SEKxxxPMI`.
- **from**: The start date in the format `yyyy-MM-dd`.
- **to**: The end date in the format `yyyy-MM-dd`.

#### Example
\`\`\`
POST /getObservation?seriesIdList=AUD,BRL&from=2023-09-01&to=2023-10-01
\`\`\`

### `/mergeData` (GET)
Merges downloaded data into a predefined Excel sheet.

#### Example
\`\`\`
GET /mergeData
\`\`\`

### `/ExcelToJson` (GET)
Converts the downloaded data saved in Excel into JSON format.

#### Example
\`\`\`
GET /ExcelToJson
\`\`\`

## Code Explanation

### `RiksbankenApiService`
A Spring service that contains methods for data fetching and transformation.

#### `fetchObservations(...)`
Fetches data from the Riksbank API based on given series IDs and a date range, then calls `jsonToExcel` to save the data to an Excel file.

#### `jsonToExcel(...)`
Takes fetched data and saves it in an Excel format. The Excel sheet will have a date column and one column for each currency.

#### `mergeData()`
Reads data from `riksbankens_kurser.xlsx` and merges it with data in `DatadumpDemo.xlsx`.

#### `ExcelToJson()`
Converts the Excel sheet to JSON format and saves it as `output.json`.

## Installation and Running
- Clone the repository.
- Open the project in your IDE and resolve the Maven dependencies.
- Run `RiksbankenController.java` to start the application.
- Use Postman or any HTTP client to test the endpoints.

## Dependencies
- Spring Boot
- Jackson for JSON handling
- Apache POI for Excel operations
