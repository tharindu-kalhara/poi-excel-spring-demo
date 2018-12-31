package com.example.poi;

import com.example.poi.dto.Customer;
import com.example.poi.help.ExcelGenerator;
import com.example.poi.repo.CustomerRepository;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.util.Date;
import java.util.List;

@RestController
@RequestMapping("/api/customers")
public class CustomerExcelDownloadRestAPI {

    @Autowired
    private CustomerRepository customerRepository;

    @GetMapping(value = "/download/customers.xlsx")
    public ResponseEntity<InputStreamResource> excelCustomersReport() throws IOException {
        List<Customer> customers = (List<Customer>) customerRepository.findAll();

        ByteArrayInputStream in = ExcelGenerator.customersToExcel(customers);
        // return IOUtils.toByteArray(in);

        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=customers.xlsx");

        return ResponseEntity
                .ok()
                .headers(headers)
                .body(new InputStreamResource(in));
    }

    @GetMapping(value = "update")
    public ResponseEntity<InputStreamResource> updateReport() throws IOException, InvalidFormatException {
        FileInputStream inputStream = new FileInputStream(new File("src/main/resources/Settlement_List_31_12_2018.xlsx"));

        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Object[][] bookData = {
                {"The Passionate Programmer", "Chad Fowler", 16},
                {"Software Craftmanship", "Pete McBreen", 26},
                {"The Art of Agile Development", "James Shore", 32},
                {"Continuous Delivery", "Jez Humble", 41},
        };
//        int rowCount = sheet.getLastRowNum();
        int rowCount = 21;
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeight((short) (9 * 20));
        font.setBold(false);
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);

        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            Cell cell = row.createCell(columnCount);
            cell.setCellValue(rowCount);
            cell.setCellStyle(style);

            for (Object field : aBook) {
                cell = row.createCell(++columnCount);
                cell.setCellStyle(style);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }


        }
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);


        inputStream.close();
//        FileOutputStream outputStream = new FileOutputStream("src/main/resources/new-customers.xlsx");
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=customers.xlsx");
        return ResponseEntity
                .ok()
                .headers(headers)
                .body(new InputStreamResource(new ByteArrayInputStream(outputStream.toByteArray())));
    }
}
