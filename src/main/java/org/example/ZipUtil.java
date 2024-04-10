package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipUtil {
        //create a zip file
        public static void zip(String zipFilePath, String... filePaths) throws IOException {
            try (ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(zipFilePath))) {
                for (String filePath : filePaths) {
                    File fileToZip = new File(filePath);
                    FileInputStream fileInputStream = new FileInputStream(fileToZip);
                    ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
                    zipOut.putNextEntry(zipEntry);
                    byte[] bytes = new byte[1024];
                    int length;
                    while ((length = fileInputStream.read(bytes)) >= 0) {
                        zipOut.write(bytes, 0, length);
                    }
                    fileInputStream.close();
                    zipOut.closeEntry();
                }
            }
        }
}
