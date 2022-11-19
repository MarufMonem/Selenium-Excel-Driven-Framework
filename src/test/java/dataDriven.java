import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class dataDriven {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("X:\\Self improvement\\Selenium Udemy\\Code\\ExcelDriven\\ExcelDrivenData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        
        int sheetCount = workbook.getNumberOfSheets();
        for(int i=0; i<sheetCount; i++){
            if(workbook.getSheetName(i).equalsIgnoreCase("sheet1")) {
                XSSFSheet sheet = workbook.getSheetAt(i); //has all the rows
                Iterator<Row> rows = sheet.iterator(); //iterates through each and every row
                Row firstRow = rows.next(); //initially goes to the first row
                Iterator<Cell> firstRowCellData = firstRow.cellIterator();
                int k =0;
                int columnNumber=0;
                while(firstRowCellData.hasNext()){
                    Cell value = firstRowCellData.next(); //goes to the first column
                    if(value.getStringCellValue().equalsIgnoreCase("Testcases")){
//                        desired column
                        columnNumber =k;
                    }
                    k++; // helps us to get column value
                }
                System.out.println("column number for Testcase: " + columnNumber);
//                break;
            }
        }
    }
}
