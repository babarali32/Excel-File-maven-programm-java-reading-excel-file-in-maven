package excelFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class excelDemo04 {
    public static void main(String[] args) throws IOException {
        String path = "Files/excel.xlsx"; // Verify the correct file path
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = xssfWorkbook.getSheet("babarsheet"); // Verify the sheet name

        int numberOfRows = sheet.getPhysicalNumberOfRows();
        System.out.println("Number of rows: " + numberOfRows);

        for (int i = 0; i < numberOfRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) { // Check if row is not null
                int numberOfCells = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfCells; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) { // Check if cell is not null
                        System.out.print(cell + " ");
                    } else {
                        System.out.print("Empty Cell ");
                    }
                }
                System.out.println();
            }
        }
    }
}

