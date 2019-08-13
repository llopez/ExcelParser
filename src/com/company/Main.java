package com.company;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello");

        try {

            Biff8EncryptionKey.setCurrentUserPassword("TabulaReporting!");
            FileInputStream input = new FileInputStream("/home/luigi/Desktop/test.xls");
            HSSFWorkbook hwb = new HSSFWorkbook(input);

            HSSFSheet sheet = hwb.getSheetAt(0);
            HSSFRow row = sheet.getRow(0);
            String value1 = row.getCell(0).getStringCellValue();
            String value2 = row.getCell(1).getStringCellValue();
            System.out.println(value1);
            System.out.println(value2);
            Biff8EncryptionKey.setCurrentUserPassword(null);

            FileOutputStream out = new FileOutputStream(new File("/home/luigi/Desktop/test-unprotected.xls"));

            hwb.write(out);
            hwb.close();
            out.close();


        } catch (Exception ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        } finally {
            Biff8EncryptionKey.setCurrentUserPassword(null);
        }
    }
}
