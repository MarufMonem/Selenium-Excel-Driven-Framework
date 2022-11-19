import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class dataDriven {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("X:\\Self improvement\\Selenium Udemy\\Code\\ExcelDriven\\ExcelDrivenData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);

    }
}
