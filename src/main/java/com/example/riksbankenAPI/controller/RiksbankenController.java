package com.example.riksbankenAPI.controller;


import com.example.riksbankenAPI.service.RiksbankenApiService;
import com.fasterxml.jackson.databind.JsonNode;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.List;

@RestController
public class RiksbankenController {

    //AUD BRL CAD CHF CNY CZK DKK EUR GBP HKD HUF IDR INR ISK JPY KRW MAD MXN NOK NZD PLN SAR SGD THB TRY USD ZAR
 //"SEKEURPMI","SEKCADPMI", "SEKAUDPMI");
  // from = "2023-09-01";
 //  to = "2023-10-01";

    @Autowired
    RiksbankenApiService riksbankenApiService;

    @PostMapping("/getObservation")
    public JsonNode getObservation(@RequestParam List<String> seriesIdList, @RequestParam String from, @RequestParam String to) {

        riksbankenApiService.fetchObservations(seriesIdList,from,to);

        return null;
    }

    @GetMapping("/mergeData")
    public String mergeData() throws IOException {

        riksbankenApiService.copyFile();

        riksbankenApiService.ExcelToJson();

        return "merged data - tester";
    }
}
