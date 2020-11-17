Apache POI integration with Spring Boot MVC
===========================================

This PoC shows how Apache POI can be integrated with SpringBoot MVC to generate a service which reads/writes excel files.

Writing excel files
-------------------

An excel file is composed of sheets which contains rows of cells, therefore the input for the writing service would look
like this:

```json
{
  "fileName": "File Name",
  "sheets": [
    {
      "title": "Sheet one",
      "rows": [
        {
          "cells": [
            {
              "type": "STRING",
              "value": "Name"
            },
            {
              "type": "STRING",
              "value": "Age"
            }
          ]
        },
        {
          "cells": [
            {
              "type": "STRING",
              "value": "Jhon Doe"
            },
            {
              "type": "INT",
              "value": "20"
            }
          ]
        }
      ]
    }
  ]
}
```

This data model is represented by the `com.example.poi.dto.ExcelDTO` class.

In order to generate the file, first we must instance an XSSFWorkbook that will represent the file contents, and create 
the specified sheets:

```java
var workbook = new XSSFWorkbook();

body.getSheets()
        .forEach(s -> buildSheet(s, workbook));
```

For building each sheet, we use the `workbook.createSheet(String)` method which returns an `XSSFSheet` instance and
add all the rows from the input with the `sheet.createRow(int)`:

```java
private void buildSheet(SheetDTO dto, XSSFWorkbook workbook) {
    var sheet = workbook.createSheet(dto.getTitle());

    for (int i = 0; i < dto.getRows().size(); i++) {
        var rowDTO = dto.getRows().get(i);
        var row = sheet.createRow(i);

        buildRow(rowDTO, row);
    }
}
```

Then, for each row we use the `row.createCell(int)` method to create a cell in the row and assign its values:

```java
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
```

In order to retrieve the generated file as an HTTP file download, we return an InputStreamResource from the controller endpoint
by writing the contents of the workbook to a byte array:

```java
var outputStream = new ByteArrayOutputStream();
workbook.write(outputStream);
workbook.close();

return ResponseEntity.ok()
        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + body.getFileName() + ".xlsx")
        .contentType(MediaType.APPLICATION_OCTET_STREAM)
        .body(new InputStreamResource(new ByteArrayInputStream(outputStream.toByteArray())));
```