package com.example.riksbankenAPI.controller;


import com.example.riksbankenAPI.service.RiksbankenApiService;
import com.fasterxml.jackson.databind.JsonNode;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import java.util.Arrays;
import java.util.List;

@RestController
public class RiksbankenController {

    //AUD BRL CAD CHF CNY CZK DKK EUR GBP HKD HUF IDR INR ISK JPY KRW MAD MXN NOK NZD PLN SAR SGD THB TRY USD ZAR
    List<String> seriesIdList = Arrays.asList("SEKEURPMI","SEKCADPMI", "SEKAUDPMI");
    String from = "2023-09-01";
    String to = "2023-10-01";

    @Autowired
    RiksbankenApiService riksbankenApiServiceTest;

    @PostMapping("/getObservation")
    public JsonNode getObservation(@RequestParam List<String> seriesIdList, @RequestParam String from, @RequestParam String to) {

        riksbankenApiServiceTest.fetchObservations(seriesIdList,from,to);

        return null;
    }

}
