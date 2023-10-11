package com.example.riksbankenAPI.service;

import org.springframework.stereotype.Service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import java.net.HttpURLConnection;
import java.net.URL;

import java.util.ArrayList;
import java.util.List;

@Service
public class RiksbankenApiService {

    public List<JsonNode> fetchObservations(List<String> seriesIds, String from, String to) {
        List<JsonNode> jsonNodes = new ArrayList<>();
        ObjectMapper objectMapper = new ObjectMapper();

        for (String seriesId : seriesIds) {
            String urlString = "https://api-test.riksbank.se/swea/v1/Observations/" + seriesId + "/" + from + "/" + to;

            try (BufferedReader in = new BufferedReader(new InputStreamReader(new URL(urlString).openStream()))) {
                StringBuilder response = new StringBuilder();
                String inputLine;

                while ((inputLine = in.readLine()) != null) {
                    response.append(inputLine);
                }

                JsonNode jsonData = objectMapper.readTree(response.toString());
                jsonNodes.add(jsonData);

            } catch (IOException e) {
                e.printStackTrace();
                return null;
            }
        }

        jsonToExcel(jsonNodes, seriesIds);
        return jsonNodes;
    }

    public void jsonToExcel(List<JsonNode> jsonNodes, List<String> seriesIds) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); FileOutputStream outputStream = new FileOutputStream("riksbankens_kurser.xlsx")) {
            XSSFSheet sheet = workbook.createSheet("Data Details");
            Row headerRow = sheet.createRow(0);

            headerRow.createCell(0).setCellValue("CurrencyID");
            headerRow.createCell(1).setCellValue("Date");
            headerRow.createCell(2).setCellValue("Value");

            int rowNum = 1;

            for (int i = 0; i < jsonNodes.size(); i++) {
                JsonNode rootNode = jsonNodes.get(i);
                String seriesId = seriesIds.get(i);

                for (JsonNode dataObject : rootNode) {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(seriesId);
                    row.createCell(1).setCellValue(dataObject.get("date").asText());
                    row.createCell(2).setCellValue(dataObject.get("value").asDouble());
                }
            }

            workbook.write(outputStream);
            System.out.println("Excel file generated");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
