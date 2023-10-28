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


    @Autowired
    RiksbankenApiService riksbankenApiService;

    @PostMapping("/getObservation")
    public String getObservation(@RequestParam List<String> seriesIdList, @RequestParam String from, @RequestParam String to) {
        riksbankenApiService.fetchObservations(seriesIdList, from, to);
        return "Downloaded data from riksbankenAPI.";
    }

    @GetMapping("/mergeData")
    public String mergeData() throws IOException {
        riksbankenApiService.mergeData();
        return "merged data.";
    }

    @GetMapping("/ExcelToJson")
    public JsonNode ExcelToJson() throws IOException {
        return riksbankenApiService.ExcelToJson();
    }
}
