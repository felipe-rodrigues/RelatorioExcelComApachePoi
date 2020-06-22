package com.reports.demo;

import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;

@RestController
public class ReportController {


    @Autowired
    ResourceLoader resourceLoader;

    @RequestMapping("/")
    String WriteHello(){
        return "hello world";
    }

    @RequestMapping("/export")
    public ResponseEntity<String> exportExcelGrafico(HttpServletResponse response){

        Resource resource = resourceLoader.getResource("classpath:gastos.xlsx");

        try {
            InputStream inputStream = resource.getInputStream();

            XSSFWorkbook book = getExcelFile(inputStream);

        } catch (IOException e) {
            e.printStackTrace();
        }

        return new ResponseEntity<String>(HttpStatus.OK);
    }

    private XSSFWorkbook getExcelFile(InputStream inputStream) {

        try{
            XSSFWorkbook book = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = book.getSheetAt(0);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);

            XSSFChart chart = drawing.createChart(anchor);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
