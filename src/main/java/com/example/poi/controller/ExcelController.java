package com.example.poi.controller;

import com.example.poi.dto.ExcelDTO;
import com.example.poi.dto.RowDTO;
import com.example.poi.dto.SheetDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;

@Slf4j
@RestController
public class ExcelController {
    @PostMapping
    public ResponseEntity<InputStreamResource> download(@RequestBody ExcelDTO body) throws Exception {
        var workbook = new XSSFWorkbook();

        body.getSheets()
                .forEach(s -> buildSheet(s, workbook));

        var outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + body.getFileName() + ".xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(new InputStreamResource(new ByteArrayInputStream(outputStream.toByteArray())));
    }

    private void buildSheet(SheetDTO dto, XSSFWorkbook workbook) {
        var sheet = workbook.createSheet(dto.getTitle());

        for (int i = 0; i < dto.getRows().size(); i++) {
            var rowDTO = dto.getRows().get(i);
            var row = sheet.createRow(i);

            buildRow(rowDTO, row);
        }
    }

    private void buildRow(RowDTO rowDTO, XSSFRow row) {
        for (int j = 0; j < rowDTO.getCells().size(); j++) {
            var cellDTO = rowDTO.getCells().get(j);
            var cell = row.createCell(j);

            switch (cellDTO.getType()) {
                case STRING:
                    cell.setCellValue(cellDTO.getValue());
                    break;

                case DOUBLE:
                    cell.setCellValue(Double.parseDouble(cellDTO.getValue()));
                    break;

                case INT:
                    cell.setCellValue(Integer.parseInt(cellDTO.getValue()));
                    break;

                case BOOLEAN:
                    cell.setCellValue(Boolean.parseBoolean(cellDTO.getValue()));
                    break;
            }
        }
    }
}
