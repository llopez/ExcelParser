package com.company;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) {
        String inputPath  = args[0];
        String outputPath = args[1];
        String password = args[2];
        try {
            Biff8EncryptionKey.setCurrentUserPassword(password);
            FileInputStream input = new FileInputStream(inputPath);
            HSSFWorkbook workbook = new HSSFWorkbook(input);
            Biff8EncryptionKey.setCurrentUserPassword(null);
            FileOutputStream out = new FileOutputStream(new File(outputPath));
            workbook.write(out);
            workbook.close();
            out.close();
        } catch (Exception ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        } finally {
            Biff8EncryptionKey.setCurrentUserPassword(null);
        }
    }
}
