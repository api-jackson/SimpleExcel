package org.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;

@Slf4j
public class ExcelImportUtils {



    public static Workbook determineWorkbook(File file) {
        Workbook workbook = null;
        try {
            String fileName = file.getName();
            int suffixIndex = fileName.lastIndexOf(".");
            String suffixName = fileName.substring(suffixIndex + 1);
            if ("xlsx".equals(suffixName)) {
                workbook = new XSSFWorkbook(file);
            } else if ("xls".equals(suffixName)) {
                FileInputStream fis = new FileInputStream(file);
                workbook = new HSSFWorkbook(fis);
            } else {
                workbook = new SXSSFWorkbook(new XSSFWorkbook(file));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
        return workbook;
    }





    /******************** 导入处理方法  start ********************/


    public static <T> List<T> getObjectListFromExcel(Workbook workbook, Class<T> klass) {
        return getObjectListFromExcel(workbook, klass, "");
    }

    public static <T> List<T> getObjectListFromExcel(Workbook workbook, Class<T> klass, String scene) {
        List<Field> importFieldList = FieldUtils.getFieldsListWithAnnotation(klass, ExcelImport.class);
        Map<ExcelImport, Field> excelImportFieldMap = filterExcelImportScene(importFieldList, scene);
        List<Field> fieldList = new ArrayList<>();
        excelImportFieldMap.forEach((annotation, field) -> fieldList.add(field));
        List<Map<String, Object>> mapListFromExcel = getMapListFromExcel(workbook, fieldList);
        List<T> objectList = getObjectList(mapListFromExcel, klass, fieldList);
        return objectList;
    }

    private static Map<ExcelImport, Field> filterExcelImportScene(List<Field> fieldList, String sceneParam) {
        Map<ExcelImport, Field> sceneFieldMap = new TreeMap<>(new Comparator<ExcelImport>() {
            @Override
            public int compare(ExcelImport o1, ExcelImport o2) {
                return o1.order()- o2.order();
            }
        });
        for (Field field : fieldList) {  // Field Loop
            ExcelImport[] excelImportArray = field.getAnnotationsByType(ExcelImport.class);
            excelImportAnnotationLoop:
            for (ExcelImport excelImport : excelImportArray) {  // Annotation ExcelExport Loop
                for (String sceneInField : excelImport.scene()) {  //  ExcelExport.scene Array Loop
                    if (sceneInField.equals(sceneParam)) {
                        sceneFieldMap.put(excelImport, field);
                        break excelImportAnnotationLoop;
                    }
                }
            }
        }
        return sceneFieldMap;
    }

    private static List<Map<String, Object>> getMapListFromExcel(Workbook workbook, List<Field> fieldList) {
        Sheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        List<Map<String, Object>> objectMapList = new ArrayList<>();
        for (int rowi = 1; rowi<=lastRowNum; rowi++) {
            Map<String, Object> excelValueMap = new HashMap<>();
            Row row = sheet.getRow(rowi);
            for (int columnj = 0; columnj < fieldList.size(); columnj++) {
                Cell cell = row.getCell(columnj);
                setFieldValue(cell, fieldList.get(columnj), excelValueMap);
            }
            objectMapList.add(excelValueMap);
        }
        return objectMapList;
    }

    private static void setFieldValue(Cell cell, Field field, Map<String, Object> excelValueMap) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            excelValueMap.put(field.getName(), null);
            return;
        }
        Class<?> type = field.getType();

        CellType cellType = cell.getCellType();

        boolean isDoubleType = type.isAssignableFrom(Double.class) || type.isAssignableFrom(double.class);
        boolean isIntergerType = type.isAssignableFrom(Integer.class) || type.isAssignableFrom(int.class);
        boolean isLongType = type.isAssignableFrom(Long.class) || type.isAssignableFrom(long.class);
        boolean isBigDecimalType = type.isAssignableFrom(BigDecimal.class);

        boolean isNumericType = isDoubleType || isIntergerType || isLongType || isBigDecimalType;

        if (type.isAssignableFrom(Date.class)
                || type.isAssignableFrom(LocalDate.class)
                || type.isAssignableFrom(LocalDateTime.class)) {

            Date dateValue = getDate(cell);

            if (type.isAssignableFrom(Date.class)) {
                excelValueMap.put(field.getName(), dateValue);
            } else if (type.isAssignableFrom(LocalDate.class)) {
                LocalDate localDate = dateValue.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                excelValueMap.put(field.getName(), localDate);
            } else if (type.isAssignableFrom(LocalDateTime.class)) {
                LocalDateTime localDateTime = LocalDateTime.ofInstant(dateValue.toInstant(), ZoneId.systemDefault());
                excelValueMap.put(field.getName(), localDateTime);
            }

        } else if (type.isAssignableFrom(Boolean.class) || type.isAssignableFrom(boolean.class)) {
            Boolean booleanValue = null;
            if (cell.getCellType() == CellType.STRING) {
                booleanValue = Boolean.parseBoolean(cell.getStringCellValue());
            } else if (cell.getCellType() == CellType.BOOLEAN) {
                booleanValue = cell.getBooleanCellValue();

            }

            excelValueMap.put(field.getName(), booleanValue);
        } else if (isNumericType) {
            Double numericValue = null;
            if (cell.getCellType() == CellType.STRING) {
                numericValue = Double.parseDouble(cell.getStringCellValue());
            } else if (cell.getCellType() == CellType.NUMERIC){
                numericValue = cell.getNumericCellValue();
            }


            if (type.isAssignableFrom(Double.class) || type.isAssignableFrom(double.class)) {
                excelValueMap.put(field.getName(), numericValue);
            } else if (type.isAssignableFrom(Integer.class) || type.isAssignableFrom(int.class)) {
                Integer intCellValue = new Double(numericValue).intValue();
                excelValueMap.put(field.getName(), intCellValue);
            } else if (type.isAssignableFrom(Long.class) || type.isAssignableFrom(long.class)) {
                Long longValue = new Double(numericValue).longValue();
                excelValueMap.put(field.getName(), longValue);
            } else if (type.isAssignableFrom(BigDecimal.class)) {
                BigDecimal decimalValue = BigDecimal.valueOf(numericValue);
                excelValueMap.put(field.getName(), decimalValue);
            }

        } else if (type.isAssignableFrom(String.class)) {
            String stringValue = StringUtils.EMPTY;
            if (cell.getCellType() == CellType.NUMERIC) {
                double numericCellValue = cell.getNumericCellValue();
                stringValue = Double.valueOf(numericCellValue).toString();
            } else if (cell.getCellType() == CellType.STRING) {
                stringValue = cell.getStringCellValue();
            }
            excelValueMap.put(field.getName(), stringValue);
        } else {
            String cellValue = cell.getStringCellValue();
            excelValueMap.put(field.getName(), cellValue);
        }
    }


    private static <T> List<T> getObjectList(List<Map<String, Object>> objectMapList, Class<T> klass, List<Field> fieldList) {
        List<T> objectList = new ArrayList<>();
        try {
            for (Map<String, Object> objectMap : objectMapList) {
                Constructor<T> constructor = klass.getConstructor();
                T t = constructor.newInstance();
                for (Field field : fieldList) {
                    field.setAccessible(true);
                    field.set(t, objectMap.get(field.getName()));
                }
                objectList.add(t);
            }
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        } catch (InvocationTargetException e) {
            throw new RuntimeException(e);
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        return objectList;
    }

    private static Date getDate(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getDateCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue();
            try {
                return DateUtils.parseDate(cellValue, "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss");
            } catch (ParseException e) {
                throw new RuntimeException(e);
            }
        }
        return cell.getDateCellValue();

    }


    /******************** 导入处理方法  end ********************/


}


