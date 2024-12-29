package org.utils.extra;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * @author Jackson
 * @date 2024/11/30
 * @description
 */
public class WorkbookHandler {

    public static String saveExcelToFile(Workbook workbook, String filename) throws IOException {
        int surfixIndex = filename.lastIndexOf(".");
        String dateString = DateTimeUtils.date2String(new Date(), DateTimeUtils.PATTERN_DATETIME_MINI);
        String nowFilename = new String(filename.substring(0, surfixIndex) + "_" + dateString + "." + filename.substring(surfixIndex + 1));
        File file = new File(nowFilename);
        file.createNewFile();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw e;
        }
        return file.getAbsolutePath();
    }

}
