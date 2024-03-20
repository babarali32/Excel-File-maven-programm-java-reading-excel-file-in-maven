package excelFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class excelMap {
        public static void main (String[]args) throws IOException {
            String path = "Files/excelmap.xlsx";
            FileInputStream fileInputStream = new FileInputStream(path);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            int numberofrows = sheet.getPhysicalNumberOfRows();
            ArrayList<Map<String,String>> arrayList=new ArrayList<>();
            Row zero=sheet.getRow(0);
            for (int i = 1; i < numberofrows; i++) {
                Row allrows = sheet.getRow(i);
                Map<String, String> data = new LinkedHashMap<>();
                if (allrows!=null){
                data.put(zero.getCell(0).toString(), allrows.getCell(0).toString());
                data.put(zero.getCell(1).toString(), allrows.getCell(1).toString());
                data.put(zero.getCell(2).toString(), allrows.getCell(2).toString());
                data.put(zero.getCell(3).toString(), allrows.getCell(3).toString());

                    arrayList.add(data);

            }

            }
              System.out.println(arrayList);

            }
        }
