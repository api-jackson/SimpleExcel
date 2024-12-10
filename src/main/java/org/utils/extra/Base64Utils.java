package org.utils.extra;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.Date;

public class Base64Utils {

    public static File base64StringToFile(String base64String, String filename) throws IOException {
        int surfixIndex = filename.lastIndexOf(".");
        String dateString = DateTimeUtils.date2String(new Date(), DateTimeUtils.PATTERN_DATETIME_MINI);
        String nowFilename = new String(filename.substring(0, surfixIndex)+"_"+dateString+"."+filename.substring(surfixIndex+1));
        File file = new File(nowFilename);
        file.createNewFile();
        byte[] decodeBytes = Base64.getDecoder().decode(base64String);
        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(decodeBytes);
            fos.flush();
        }

        return file;


    }

}
