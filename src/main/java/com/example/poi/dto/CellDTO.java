package com.example.poi.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CellDTO {
    public enum CellType {
        STRING,
        DOUBLE,
        INT,
        BOOLEAN
    }

    private String value;
    private CellType type;
}
