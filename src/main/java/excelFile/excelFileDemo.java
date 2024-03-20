package excelFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class excelFileDemo {
    public static void main(String[] args) throws IOException {
        String path = "Files/excel.xlsx";
        FileInputStream fileInputStream = new FileInputStream(path);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("babarsheet");
        int numberofrows=sheet.getPhysicalNumberOfRows();
        System.out.println(numberofrows);
        for (int i = 0; i < numberofrows; i++) {
            Row allrows=sheet.getRow(i);
            if (allrows !=null){
              int allcels=allrows.getPhysicalNumberOfCells();
                for (int j = 0; j < allcels; j++) {
                   Cell cels=allrows.getCell(j);
                   if (cels!=null){
                       System.out.print(cels+" ");
                   }else {
                       System.out.println("cell is empty");
                   }
                }
            }
            System.out.println();
        }
    }
}
