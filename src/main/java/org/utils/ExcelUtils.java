package org.utils;

import com.google.common.collect.BiMap;
import com.google.common.collect.HashBiMap;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;

@Slf4j
public class ExcelUtils {



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

    /******************** 导出处理方法  start ********************/

    private static void buildTitle(Row row, List<String> titleList) {
        for (int i=0; i<titleList.size(); i++) {
            String title = titleList.get(i);
            Cell cell = row.createCell(i, CellType.STRING);
            cell.setCellValue(title);
        }
    }

    private static void getExcelBytes(List<PropertyNameType> propertyTypeMapList, List<Map<String, Object>> exportObjectMapList,
                                     List<String> titleList, Workbook workbook, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row titleRow = sheet.createRow(0);
        buildTitle(titleRow, titleList);
        for (int rownum = 0; rownum<exportObjectMapList.size(); rownum++) {  // 行处理
            Row row = sheet.createRow(rownum+1);
            Map<String, Object> rowObject = exportObjectMapList.get(rownum);
            for (int column=0; column < propertyTypeMapList.size(); column++) {  // 行中每个单元格处理
                PropertyNameType propertyNameType = propertyTypeMapList.get(column);
                Cell cell = row.createCell(column, propertyNameType.getPropertyType());
                Object exportObject = rowObject.get(propertyNameType.getPropertyName());
                if (rowObject.get(propertyNameType.getPropertyName()) == null) {  // 对象为空，跳过
                    continue;
                }
                cell.setCellValue(exportObject.toString());
            }
        }
    }



//    public static void download(HttpServletResponse response, Workbook workbook, String filename) throws IOException {
//        int surfixIndex = filename.lastIndexOf(".");
//        String dateString = DateTimeUtils.date2String(new Date(), DateTimeUtils.PATTERN_DATETIME_MINI);
//        String nowFilename = new String(filename.substring(0, surfixIndex)+"_"+dateString+"."+filename.substring(surfixIndex+1));
//        File file = new File(nowFilename);
//        file.createNewFile();
//        try (FileOutputStream fos = new FileOutputStream(file)) {
//            workbook.write(fos);
//            fos.flush();
//        } catch (Exception e) {
//            throw e;
//        }
//
//        FileInputStream inputStream = new FileInputStream(file);
//        try (OutputStream outputStream = new BufferedOutputStream(response.getOutputStream())) {
//            byte[] bytes = StreamUtils.copyToByteArray(inputStream);
//            String exportFilename = new String(nowFilename.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
//            response.addHeader(HttpHeaders.CONTENT_DISPOSITION,
//                    "attachment;filename=" + exportFilename);
//            response.addHeader("Content-length", String.valueOf(bytes.length));
//            response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
//            outputStream.write(bytes);
//            outputStream.flush();
//        } catch (Exception e) {
//            throw e;
//        } finally {
//            if (inputStream != null) {
//                inputStream.close();
//            }
//        }
//        file.delete();
//        response.getOutputStream();
//        outputStream.write(bytes);
//        response.flushBuffer();
//    }

    private static <T> Map<String, List> generateExportList(List<T> queryList, Class<T> klass, String scene) throws IllegalAccessException {
        List<Field> exportFieldList = FieldUtils.getFieldsListWithAnnotation(klass, ExcelExport.class);
        Map<ExcelExport, Field> excelExportFieldMap = filterScene(exportFieldList, scene);
        BiMap<Field, ExcelExport> excelExportFieldBiMap = HashBiMap.create(excelExportFieldMap).inverse();
        Set<Map.Entry<ExcelExport, Field>> excelExportSet = excelExportFieldMap.entrySet();

        List<String> titleList = new ArrayList<>();
        List<PropertyNameType> propertyNameTypeList = new ArrayList<>();
        List<Map<String, Object>> exportList = new ArrayList<>();
        List<Field> fieldList = new ArrayList<>();

        for (Map.Entry<ExcelExport, Field> entry : excelExportSet) {
            ExcelExport excelExport = entry.getKey();
            Field field = entry.getValue();
            fieldList.add(field);
            titleList.add(excelExport.title());
            propertyNameTypeList.add(new PropertyNameType(field.getName(), excelExport.cellType()));
        }


        for (T t : queryList) {
            Map<String, Object> objectMap = new HashMap<>();
            for (Field field : fieldList) {
                field.setAccessible(true);
                Object object = field.get(t);
                object = processBySerializeBean(object, excelExportFieldBiMap.get(field));
                objectMap.put(field.getName(), object);
            }
            exportList.add(objectMap);
        }

        Map<String, List> exportExcelListMap = new HashMap<>();
        exportExcelListMap.put("titleList", titleList);
        exportExcelListMap.put("propertyList", propertyNameTypeList);
        exportExcelListMap.put("objectList", exportList);
        return exportExcelListMap;

    }

    private static Object processBySerializeBean(Object object, ExcelExport excelExport) {
        final boolean emptyBeanName = StringUtils.isBlank(excelExport.serializeBeanName());
        final boolean emptyBeanClass = excelExport.serializeBeanName() == null || excelExport.serializeBeanClass().equals(void.class);
        if (emptyBeanName && emptyBeanClass) {  // beanName 且 beanClass 都为空
            return object;
        }
        if (!emptyBeanName) {  // beanName 不为空，优先取beanName
            return processBySerializeBeanName(object, excelExport);
        } else {  // beanClass 不为空
            return processBySerializeBeanClass(object, excelExport);
        }
    }

    private static Object processBySerializeBeanName(Object object, ExcelExport excelExport) {
        final AbstractExcelSerialize serializeBean = (AbstractExcelSerialize) SpringBeanUtils.getBean(excelExport.serializeBeanName());
        return serializeBean.getExcelObject(object);
    }

    private static Object processBySerializeBeanClass(Object object, ExcelExport excelExport) {
        final AbstractExcelSerialize serializeBean = (AbstractExcelSerialize) SpringBeanUtils.getBean(excelExport.serializeBeanClass());
        return serializeBean.getExcelObject(object);
    }

    private static Map<ExcelExport, Field> filterScene(List<Field> fieldList, String sceneParam) {
        Map<ExcelExport, Field> sceneFieldMap = new TreeMap<>(new Comparator<ExcelExport>() {
            @Override
            public int compare(ExcelExport o1, ExcelExport o2) {
                return o1.order()- o2.order();
            }
        });

        for (Field field : fieldList) {  // Field Loop
            ExcelExport[] excelExportArray = field.getAnnotationsByType(ExcelExport.class);
            excelExportAnnotationLoop:
            for (ExcelExport excelExport : excelExportArray) {  // Annotation ExcelExport Loop
                for (String sceneInField : excelExport.scene()) {  //  ExcelExport.scene Array Loop
                    if (sceneInField.equals(sceneParam)) {
                        sceneFieldMap.put(excelExport, field);
                        break excelExportAnnotationLoop;
                    }
                }
            }
        }
        return sceneFieldMap;
    }


    private static void processExcelSheetByListMap(Map<String, List> exportListMap, Workbook workbook, String sheetName) {
        List<String> titleList = exportListMap.get("titleList");
        List<PropertyNameType> propertyNameTypeList = exportListMap.get("propertyList");
        List<Map<String, Object>> objectList = exportListMap.get("objectList");
        ExcelUtils.getExcelBytes(propertyNameTypeList, objectList, titleList, workbook, sheetName);
    }

    public static <T> void exportSheet(Workbook workbook, List<T> dataList, Class<T> klass, String sheetName, String scene) {
        try {
            Map<String, List> exportListMap = ExcelUtils.generateExportList(dataList, klass, scene);
            ExcelUtils.processExcelSheetByListMap(exportListMap, workbook, sheetName);
        } catch (IllegalAccessException e) {
            log.error("导出"+sheetName+"异常", e);
        }
    }

    public static <T> void exportSheet(Workbook workbook, List<T> dataList, Class<T> klass, String sheetName) {
        exportSheet(workbook, dataList, klass, sheetName, "");
    }

    /******************** 导出处理方法  end ********************/




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

        cell.getCellType();

        boolean isDoubleType = type.isAssignableFrom(Double.class) || type.isAssignableFrom(double.class);
        boolean isIntergerType = type.isAssignableFrom(Integer.class) || type.isAssignableFrom(int.class);
        boolean isLongType = type.isAssignableFrom(Long.class) || type.isAssignableFrom(long.class);
        boolean isBigDecimalType = type.isAssignableFrom(BigDecimal.class);

        boolean isNumericType = isDoubleType || isIntergerType || isLongType || isBigDecimalType;

        if (type.isAssignableFrom(Date.class)
                || type.isAssignableFrom(LocalDate.class)
                || type.isAssignableFrom(LocalDateTime.class)) {
            Date dateValue = cell.getDateCellValue();

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
            boolean booleanValue = cell.getBooleanCellValue();
            excelValueMap.put(field.getName(), booleanValue);
        } else if (isNumericType) {

            double numericValue = cell.getNumericCellValue();

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



    /******************** 导入处理方法  end ********************/


}


