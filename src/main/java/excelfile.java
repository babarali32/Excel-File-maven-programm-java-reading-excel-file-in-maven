import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
public class excelfile {
    public static void main(String[] args) throws IOException {
        String path = "Files/excel.xlsx"; // Verify the correct file path

        FileInputStream fileInputStream = new FileInputStream(path);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = xssfWorkbook.getSheet("babarsheet"); // Verify the sheet name

        // Assuming you want to retrieve the cell at row 0 (header row) and column index 3
        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(3);

        System.out.println("Entire row: " + row); // Print the entire row
        System.out.println("Cell value: " + cell); // Print the cell value
    }
}