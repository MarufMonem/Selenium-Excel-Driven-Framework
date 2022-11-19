import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
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
                            Cell c = cellVal.next(); //storing the cell value as a cell object
//                            Checking the cell value type
                            if(c.getCellType()== CellType.STRING){
                                testData.add(c.getStringCellValue());
                            }else {
//                                This is a POI built in method to convert numerical cell values to Strings
                                testData.add(NumberToTextConverter.toText(c.getNumericCellValue()));
                            }

                        }
                    }

                }
//                break;
            }
        }
        return testData;
    }
}
