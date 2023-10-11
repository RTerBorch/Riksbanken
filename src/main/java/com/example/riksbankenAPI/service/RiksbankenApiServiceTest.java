package com.example.riksbankenAPI.service;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.util.JSONPObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.io.BufferedReader;
import java.util.ArrayList;
import java.util.List;

@Service
public class RiksbankenApiServiceTest {

    public List<JsonNode> fetchObservations(List<String> seriesIds, String from, String to) {
        List<JsonNode> jsonNodes = new ArrayList<>();

        for (String seriesId : seriesIds) {
            StringBuilder response = new StringBuilder();//
            try {
                String urlString= "https://api-test.riksbank.se/swea/v1/Observations/" + seriesId + "/" + from + "/" + to;
                URL url = new URL(urlString);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");

                BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                String inputLine;

                while ((inputLine = in.readLine()) != null) {
                    response.append(inputLine);
                }
                in.close();

                // Convert the response to a JsonNode object
                ObjectMapper objectMapper = new ObjectMapper();
                JsonNode jsonData = objectMapper.readTree(response.toString());
                jsonNodes.add(jsonData);

            } catch (Exception e) {
                e.printStackTrace();
                return null;
            }
        }

        // Call jsonToExcel with the list of JsonNode objects
        jsonToExcel(jsonNodes, seriesIds);

        return jsonNodes;
    }

    public void jsonToExcel(List<JsonNode> jsonNode, List<String> seriesId) {
        ObjectMapper om = new ObjectMapper();
        try {
            // Create a new Excel workbook and sheet
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet("Data Details");

            // Create header cells
            Row headerRow = sheet.createRow(0);

            Cell currencyHeaderCell = headerRow.createCell(0);
            currencyHeaderCell.setCellValue("CurrencyID");
            Cell dateHeaderCell = headerRow.createCell(1);
            dateHeaderCell.setCellValue("Date");
            Cell valueHeaderCell = headerRow.createCell(2);
            valueHeaderCell.setCellValue("Value");

            // Initialize the row number
            int rowNum = 1;

            System.out.println(seriesId.size() + "size");
            System.out.println("jsonNode values:");
            for (JsonNode node : jsonNode) {
                System.out.println(node.toString());
            }
            System.out.println("seriesId values:");
            for (String id : seriesId) {
                System.out.println(id);
            }

            for (int i = 0; i < jsonNode.size(); i++) {
                JsonNode dataNode = jsonNode.get(i);
                String series = seriesId.get(i);

                for (JsonNode dataObject : dataNode) {
                    Row dataRow = sheet.createRow(rowNum);

                    // Populate the data row
                    Cell currencyCell = dataRow.createCell(0);
                    currencyCell.setCellValue(series);

                    Cell dateCell = dataRow.createCell(1);
                    dateCell.setCellValue(dataObject.get("date").asText());

                    Cell valueCell = dataRow.createCell(2);
                    valueCell.setCellValue(dataObject.get("value").asDouble());

                    // Increment the row number
                    rowNum++;
                }
            }

            // Save the workbook to a file
            FileOutputStream outputStream = new FileOutputStream("Output2_data.xlsx");
            wb.write(outputStream);
            wb.close();
            System.out.println("Excel file generated");

        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
