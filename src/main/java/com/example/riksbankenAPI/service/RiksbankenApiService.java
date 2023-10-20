package com.example.riksbankenAPI.service;

// Importing necessary libraries

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

// Service annotation tells Spring Boot that this class is a service class.
@Service
public class RiksbankenApiService {

    // Method to fetch observations from the Riksbank API based on given series IDs and date range.
    public List<JsonNode> fetchObservations(List<String> seriesIds, String from, String to) {
        List<JsonNode> jsonNodes = new ArrayList<>();
        ObjectMapper objectMapper = new ObjectMapper();

        // Loop through each seriesId to construct the URL and fetch data.
        for (String seriesId : seriesIds) {
            String urlString = "https://api-test.riksbank.se/swea/v1/Observations/" + seriesId + "/" + from + "/" + to;

            try (BufferedReader in = new BufferedReader(new InputStreamReader(new URL(urlString).openStream()))) {
                StringBuilder response = new StringBuilder();
                String inputLine;

                // Read the API response line by line.
                while ((inputLine = in.readLine()) != null) {
                    response.append(inputLine);
                }

                // Convert the response to a JsonNode object.
                JsonNode jsonData = objectMapper.readTree(response.toString());
                jsonNodes.add(jsonData);

            } catch (IOException e) {
                e.printStackTrace();
                return null;
            }
        }

        // Convert the fetched JSON data to Excel format.
        jsonToExcel(jsonNodes, seriesIds);
        return jsonNodes;
    }

    // Method to convert JSON data to Excel format.
    public void jsonToExcel(List<JsonNode> jsonNodes, List<String> seriesIds) {

        try (XSSFWorkbook workbook = new XSSFWorkbook(); FileOutputStream outputStream = new FileOutputStream("src/demoWorksheet/input/riksbankens_kurser.xlsx")) {
            XSSFSheet sheet = workbook.createSheet("Data Details");

            // Create the header row with 'Date' and series IDs.
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Date");
            for (int i = 0; i < seriesIds.size(); i++) {
                headerRow.createCell(i + 1).setCellValue(seriesIds.get(i));
            }

            // Iterate over each JSON data node to populate the Excel sheet.
            for (int i = 0; i < jsonNodes.size(); i++) {
                JsonNode rootNode = jsonNodes.get(i);
                String seriesId = seriesIds.get(i);

                for (JsonNode dataObject : rootNode) {
                    String date = dataObject.get("date").asText();
                    double value = dataObject.get("value").asDouble();

                    // Fetch the row for a particular date or create a new row if not exists.
                    int rowIndex = getRowIndexOrCreate(sheet, date);
                    Row row = sheet.getRow(rowIndex);

                    // Find column index for the current series ID.
                    int colIndex = seriesIds.indexOf(seriesId) + 1;
                    Cell cell = row.createCell(colIndex, CellType.NUMERIC);
                    cell.setCellValue(value);
                }
            }

            // Write data to the Excel file.
            workbook.write(outputStream);
            System.out.println("Excel file generated");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Utility method to get row index for a particular date or create a new row if not exists.
    private int getRowIndexOrCreate(XSSFSheet sheet, String date) {
        for (Row row : sheet) {
            Cell cell = row.getCell(0);
            if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals(date)) {
                return row.getRowNum();
            }
        }
        Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
        newRow.createCell(0).setCellValue(date);
        return newRow.getRowNum();
    }

    public void mergeData2() throws IOException {
        // Load DatadumpDemo.xlsx
        FileInputStream fis1 = new FileInputStream(new File("src/demoWorksheet/input/DatadumpDemo.xlsx"));
        Workbook workbook1 = new XSSFWorkbook(fis1);
        Sheet sheet1 = workbook1.getSheetAt(0);

        // Load riksbankens_kurser.xlsx
        FileInputStream fis2 = new FileInputStream(new File("src/demoWorksheet/input/riksbankens_kurser.xlsx"));
        Workbook workbook2 = new XSSFWorkbook(fis2);
        Sheet sheet2 = workbook2.getSheetAt(0);

        // Create copy of DatadumpDemo.xlsx as merged_data.xlsx
        FileOutputStream fos = new FileOutputStream(new File("src/demoWorksheet/output/merged_data.xlsx"));
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet();

        // Copy DatadumpDemo.xlsx to merged_data.xlsx
        for (Row row : sheet1) {
            Row newRow = outputSheet.createRow(row.getRowNum());
            for (Cell cell : row) {
                Cell newCell = newRow.createCell(cell.getColumnIndex(), cell.getCellType());
                switch (cell.getCellType()) {
                    case STRING:
                        newCell.setCellValue(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(cell.getNumericCellValue());
                        break;
                    // Add more cases if required
                }
            }
        }

        // Map dates from DatadumpDemo.xlsx to row numbers for easy lookup
        HashMap<String, Integer> dateToRowNum = new HashMap<>();
        for (Row row : outputSheet) {
            Cell dateCell = row.getCell(0);
            if (dateCell != null && dateCell.getCellType() == CellType.STRING) {
                dateToRowNum.put(dateCell.getStringCellValue(), row.getRowNum());
            }
        }

        // Add headers from riksbankens_kurser.xlsx
        Row firstRowSheet2 = sheet2.getRow(0);
        Row firstRowOutput = outputSheet.getRow(0);
        int lastColumnIndex = firstRowOutput.getLastCellNum();
        for (int i = 1; i < firstRowSheet2.getLastCellNum(); i++) {
            Cell newCell = firstRowOutput.createCell(lastColumnIndex++);
            newCell.setCellValue(firstRowSheet2.getCell(i).getStringCellValue());
        }

        // Merge data
        for (Row row : sheet2) {
            Cell dateCell = row.getCell(0);
            if (dateCell != null && dateCell.getCellType() == CellType.STRING) {
                if (dateToRowNum.containsKey(dateCell.getStringCellValue())) {
                    int rowNum = dateToRowNum.get(dateCell.getStringCellValue());
                    Row targetRow = outputSheet.getRow(rowNum);
                    int colIndex = firstRowOutput.getLastCellNum() - (firstRowSheet2.getLastCellNum() - 1);
                    for (int i = 1; i < row.getLastCellNum(); i++) {
                        Cell newCell = targetRow.createCell(colIndex++);
                        switch (row.getCell(i).getCellType()) {
                            case STRING:
                                newCell.setCellValue(row.getCell(i).getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(row.getCell(i).getNumericCellValue());
                                break;
                            // Add more cases if required
                        }
                    }
                }
            }
        }

        // Save merged_data.xlsx
        outputWorkbook.write(fos);

        // Close resources
        fis1.close();
        fis2.close();
        fos.close();
        workbook1.close();
        workbook2.close();
        outputWorkbook.close();
    }

    public void mergeData() throws IOException {
        // Paths to the input files
        String path1 = "src/demoWorksheet/input/DatadumpDemo.xlsx";
        String path2 = "src/demoWorksheet/input/riksbankens_kurser.xlsx";

        // Load both workbooks
        Workbook wb1 = new XSSFWorkbook(new FileInputStream(path1));
        Workbook wb2 = new XSSFWorkbook(new FileInputStream(path2));

        // Get the sheets from both workbooks
        Sheet sheet1 = wb1.getSheetAt(0);
        Sheet sheet2 = wb2.getSheetAt(0);

        // Create a map to hold the values from the second sheet
        Map<String, Row> data = new HashMap<>();

        // Populate the map using the date as the key
        Iterator<Row> iterator = sheet2.iterator();
        iterator.next(); // Skip header
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Cell dateCell = currentRow.getCell(0);
            if (dateCell.getCellType() == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(dateCell)) {
                    String dateKey = dateCell.getDateCellValue().toString();
                    data.put(dateKey, currentRow);
                }
            }
        }

        // Merge data
        for (Row row : sheet1) {
            if (row.getRowNum() == 0) { // header
                int lastCol = row.getLastCellNum();
                Row headerRow2 = sheet2.getRow(0);
                for (int i = 1; i < headerRow2.getLastCellNum(); i++) {
                    Cell newHeader = row.createCell(lastCol - 1 + i);
                    newHeader.setCellValue(headerRow2.getCell(i).getStringCellValue());
                }
            } else {
                Cell dateCell = row.getCell(0);
                if (DateUtil.isCellDateFormatted(dateCell)) {
                    String dateKey = dateCell.getDateCellValue().toString();
                    if (data.containsKey(dateKey)) {
                        Row matchingRow = data.get(dateKey);
                        for (int i = 1; i < matchingRow.getLastCellNum(); i++) {
                            Cell cell = matchingRow.getCell(i);
                            Cell newCell = row.createCell(row.getLastCellNum());
                            newCell.setCellValue(cell.getNumericCellValue());
                        }
                    }
                }
            }
        }
// Write the result to the output file
        FileOutputStream fileOut = new FileOutputStream("src/demoWorksheet/output/merged_data.xlsx");
        wb1.write(fileOut);
        fileOut.close();

        // Close resources
        wb1.close();
        wb2.close();
    }

    public void ExcelToJson() throws IOException {
        String excelFilePath = "src/demoWorksheet/input/riksbankens_kurser.xlsx";
        Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
        Sheet sheet = workbook.getSheetAt(0);
        Gson gson = new Gson();

        List<Map<String, Object>> finalJsonList = new ArrayList<>();

        Row headerRow = sheet.getRow(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            Map<String, Object> jsonObject = new HashMap<>();
            jsonObject.put("date", row.getCell(0).toString());

            List<Map<String, Object>> list = new ArrayList<>();
            for (int i = 1; i < headerRow.getPhysicalNumberOfCells(); i++) {
                Map<String, Object> item = new HashMap<>();
                item.put("currencyName", headerRow.getCell(i).toString());
                item.put("currencyValue", Double.parseDouble(row.getCell(i).toString()));
                list.add(item);
            }
            jsonObject.put("list", list);
            finalJsonList.add(jsonObject);
        }

        String json = gson.toJson(finalJsonList);

        try (FileOutputStream fos = new FileOutputStream("src/demoWorksheet/output/output.json")) {
            byte[] jsonBytes = json.getBytes();
            fos.write(jsonBytes);
        }

        workbook.close();
    }
   public void copyFile() throws IOException {
        Files.copy(Paths.get("src/demoWorksheet/input/DatadumpDemo.xlsx"), Paths.get("src/demoWorksheet/output/DEMO.xlsx"));
    }


}
