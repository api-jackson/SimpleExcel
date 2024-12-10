package org.utils;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellType;

@Data
@AllArgsConstructor
public class PropertyNameType {
    String propertyName;

    CellType propertyType;
}