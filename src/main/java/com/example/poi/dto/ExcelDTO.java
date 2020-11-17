package com.example.poi.dto;

import lombok.*;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelDTO {
    private String fileName;
    private List<SheetDTO> sheets;
}
