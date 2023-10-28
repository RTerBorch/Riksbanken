package com.example.riksbankenAPI.service;

// Importing necessary libraries


import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@Service
public class RiksbankenApiService {
    private String dataPath = "src/demoWorksheet/input/DatadumpDemo.xlsx";
    private String currPath = "src/demoWorksheet/input/riksbankens_kurser.xlsx";


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

        try (XSSFWorkbook workbook = new XSSFWorkbook(); FileOutputStream outputStream = new FileOutputStream(currPath)) {
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

    public void mergeData() throws IOException {
        // Initialize helper objects and variables
        Map<String, Map<LocalDate, Double>> currencyMap = readExcelToCurrencyMap(currPath);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

        try (FileInputStream fis = new FileInputStream(dataPath);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream("src/demoWorksheet/output/mergedData.xlsx")) {

            Sheet sheet = workbook.getSheetAt(0);

            for (String currencyKey : currencyMap.keySet()) {
                createHeader(sheet, currencyKey);
                populateData(sheet, currencyMap.get(currencyKey), sdf);
            }

            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Finished.");
    }

    private void createHeader(Sheet sheet, String currencyKey) {
        int lastCellNum = sheet.getRow(0).getLastCellNum();
        Cell newCell = sheet.getRow(0).createCell(lastCellNum);
        newCell.setCellValue(currencyKey);
    }

    private void populateData(Sheet sheet, Map<LocalDate, Double> dateValueMap, SimpleDateFormat sdf) {
        String tmpDateString = "";
        String tmpCurValue;

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }

            boolean match = false;
            String dateStr = getDateFromCell(row.getCell(0), sdf);

            for (Map.Entry<LocalDate, Double> entry : dateValueMap.entrySet()) {
                if (dateStr != null && dateStr.equals(entry.getKey().toString())) {
                    match = true;
                    tmpCurValue = entry.getValue().toString();
                    setCellValue(row, tmpCurValue);
                    tmpDateString = entry.getKey().toString();
                }
            }
            if (!match && !tmpDateString.isEmpty()) {
                setCellValue(row, dateValueMap.get(LocalDate.parse(tmpDateString)).toString() + "*");
            }
        }
    }
    private String getDateFromCell(Cell cell, SimpleDateFormat sdf) {
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return sdf.format(cell.getDateCellValue());
        } else if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return null;
    }

    private void setCellValue(Row row, String value) {
        int lastCellNum = row.getLastCellNum();
        Cell newCell = row.createCell(lastCellNum);
        newCell.setCellValue(value);
    }


    public JsonNode ExcelToJson() throws IOException {
        String excelFilePath = currPath;
        Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
        Sheet sheet = workbook.getSheetAt(0);

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

        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode jsonNode = objectMapper.valueToTree(finalJsonList);

        try (FileOutputStream fos = new FileOutputStream("src/demoWorksheet/output/output.json")) {
            String json = objectMapper.writeValueAsString(jsonNode);
            byte[] jsonBytes = json.getBytes();
            fos.write(jsonBytes);
        }

        workbook.close();
        return jsonNode;
    }

    public Map<String, Map<LocalDate, Double>> readExcelToCurrencyMap(String filePath) {
        Map<String, Map<LocalDate, Double>> currencyMap = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            for (int i = 1; i < headerRow.getPhysicalNumberOfCells(); i++) {
                String currencyCode = headerRow.getCell(i).getStringCellValue();
                currencyMap.put(currencyCode, new HashMap<>());
            }

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                Cell dateCell = row.getCell(0);

                String dateStr;
                if (dateCell.getCellType() == CellType.STRING) {
                    dateStr = dateCell.getStringCellValue();
                } else if (dateCell.getCellType() == CellType.NUMERIC) {
                    Date javaDate = dateCell.getDateCellValue();
                    dateStr = new SimpleDateFormat("yyyy-MM-dd").format(javaDate);
                } else {
                    continue;
                }

                LocalDate date = LocalDate.parse(dateStr, DateTimeFormatter.ofPattern("yyyy-MM-dd"));

                for (int j = 1; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell rateCell = row.getCell(j);
                    double rate;
                    if (rateCell.getCellType() == CellType.STRING) {
                        String rateStr = rateCell.getStringCellValue().replace(",", ".");
                        rate = Double.parseDouble(rateStr);
                    } else if (rateCell.getCellType() == CellType.NUMERIC) {
                        rate = rateCell.getNumericCellValue();
                    } else {
                        continue;
                    }

                    String currencyCode = headerRow.getCell(j).getStringCellValue();
                    currencyMap.get(currencyCode).put(date, rate);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return currencyMap;
    }

}
