package excelFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;
public class Demo5 {

    public static void main(String[] args) throws IOException {
        String path = "Files/excelmap.xlsx";
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int numberofrows = sheet.getPhysicalNumberOfRows();
        ArrayList<Map<String, String>> arrayList = new ArrayList<>();
        Row zero = sheet.getRow(0);
        for (int i = 1; i < numberofrows; i++) {
            Row allrows = sheet.getRow(i);
            Map<String, String> data = new LinkedHashMap<>();
            if (allrows != null) {
                int numberOfCells = allrows.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfCells; j++) {
                    Cell emtpycell = allrows.getCell(i);
                    if (emtpycell != null) {
                        data.put(zero.getCell(j).toString(), allrows.getCell(j).toString());
                    }
                }
                arrayList.add(data);
            }
        }
        System.out.println(arrayList);
    }
}
