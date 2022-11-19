import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class dataDriven {

    public static ArrayList<String> getData(String testCaseName) throws IOException{
        ArrayList<String> testData = new ArrayList<String>();

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
                        columnNumber =k; //with this we know which column has the test case names
                    }
                    k++; // helps us to get column value
                }
                System.out.println("column number for Testcase: " + columnNumber);
                while (rows.hasNext()){
                    Row r  = rows.next(); //go to the first row

                    // get me the cell value only at the Row X ColumnNumber value
                    if(r.getCell(columnNumber).getStringCellValue().equalsIgnoreCase
                            (testCaseName)){
//                        pull all the data of the row:
                        Iterator<Cell> cellVal = r.cellIterator();
                        while(cellVal.hasNext()){
                            String data = cellVal.next().getStringCellValue();
                            System.out.println(data);
                            testData.add(data);
                        }
                    }

                }
//                break;
            }
        }
        return testData;
    }
    public static void main(String[] args) throws IOException {
        ArrayList<String> data= getData("Purchase");
    }
}
