package Zadatak1;

import com.github.javafaker.Faker;
import com.sun.corba.se.impl.presentation.rmi.IDLTypeException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadFile {
    public static void main(String[] args) throws IOException {

        readExcel("domaci22.xlsx");
        try {
            writeExcel("test.xlsx");
        } catch (FileNotFoundException e) {
            //throw new RuntimeException(e);
        }

        Faker faker = new Faker();
        try {
            writeFaker(faker);
        } catch (FileNotFoundException e) {
        }
    }
    private static void writeExcel(String filename) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        for (int i = 0; i < 1; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 1; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Jelena");
            }
            for (int j = 1; j < 2; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Vasilijevic");
            }
        }
        for (int i = 1; i < 2; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 1; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Pera");
            }
            for (int j = 1; j < 2; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Peric");
            }

        }
        FileOutputStream fileOutputStream = new FileOutputStream("test.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private static void writeFaker(Faker faker) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        for (int i = 2; i < 10; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 1; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(faker.name().firstName());
            }
            for (int j = 1; j < 2; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(faker.name().lastName());
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream("test.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }



    public static void readExcel(String path) {
        try {
            FileInputStream inputStream = new FileInputStream(new File("domaci22.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for (int j = 0; j < 2; j++) {

                XSSFRow row = sheet.getRow(j);

                for (int i = 0; i < 2; i++) {
                    XSSFCell cell = row.getCell(i);
                    String name = cell.getStringCellValue();
                    System.out.println(name);
                }
            }
        } catch (FileNotFoundException ex) {
            System.out.println("FIleNotFound.class");
        } catch (IOException e) {
            // e.printStackTrace();
        }
    }

}