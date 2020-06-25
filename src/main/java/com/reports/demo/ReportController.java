package com.reports.demo;

import org.apache.commons.math3.optim.nonlinear.scalar.noderiv.BOBYQAOptimizer;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
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

            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition",
                    "attachment; filename=\""
                            + "Gerando Gráfico excel"
                            + ".xlsx\"");

            book.write(response.getOutputStream());
            return new ResponseEntity<String>(HttpStatus.OK);

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
            XSSFRow row = sheet.getRow(1);
            chart.setTitleText(row.getCell(0).getStringCellValue());
            chart.setTitleOverlay(false);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            XDDFDataSource<String> categories =  XDDFDataSourcesFactory.fromStringCellRange(sheet,
                    new CellRangeAddress(0, 0, 1, 6));
            XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(1, 1, 1, 6));

            XDDFChartData data = chart.createData(ChartTypes.PIE,null,null);
            data.setVaryColors(true);
            data.addSeries(categories, val);
            chart.plot(data);

            return book;

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

}
