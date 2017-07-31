      ppackage webScrapperasd;
import java.util.*; 
import java.lang.*; 
import  java.io.*;  
import org.apache.poi.xssf.usermodel.XSSFCell; 
import org.apache.poi.xssf.usermodel.XSSFRow; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
import org.openqa.selenium.By; 
import org.openqa.selenium.WebDriver; 
import org.openqa.selenium.WebElement; 
import org.openqa.selenium.firefox.FirefoxDriver;  





    import java.io.FileInputStream;
    import java.io.FileNotFoundException;
    import java.io.FileOutputStream;
    import java.io.IOException;

    import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
    import org.apache.poi.ss.usermodel.Cell;
    import org.apache.poi.ss.usermodel.Row;
    import org.apache.poi.ss.usermodel.Sheet;
    import org.apache.poi.ss.usermodel.Workbook;
    import org.apache.poi.ss.usermodel.WorkbookFactory;

    public class WebTableTOSpreedsheet {

        private static final Date AllElements = null;


        public String getExcelData(String sheetName , int rowNum , int colNum) throws InvalidFormatException, IOException{

          WebDriver driver = new FirefoxDriver();
            driver.get("https://en.wikipedia.org/wiki/Potato");


            //get only table tex





            List<WebElement> allElements = driver.findElements(By.xpath(".//*[@id='mw-content-text']/div/table[4]/tbody[1]/tr//td[2]")); 

            for (WebElement element: allElements) {
                  System.out.println(element.getText());
            }




        String filePath = "C:\\Users\\target.xlsx";



              FileInputStream fis = new FileInputStream(filePath);
              Workbook wb = WorkbookFactory.create(fis);
              Sheet sh = wb.getSheet(sheetName);    
              Row row = sh.getRow(rowNum);
              String data = row.getCell(colNum).getStringCellValue();
              return data;
        }

        public int getRowCount(String sheetName) throws InvalidFormatException, IOException{

              FileInputStream fis = new FileInputStream("C:\\Users\\target.xlsx");
              Workbook wb = WorkbookFactory.create(fis);
              Sheet sh = wb.getSheet(sheetName);
              int rowCount = sh.getLastRowNum()+1;
            return rowCount;
        }

        public void setExcelData(String sheetName,int rowNum,int colNum,String data) throws InvalidFormatException, IOException{
              FileInputStream fis = new FileInputStream("C:\\Users\\target.xlsx");
              Workbook wb = WorkbookFactory.create(fis);
              Sheet sh = wb.getSheet(sheetName);
              Row row = sh.getRow(rowNum);
              Cell element = row.createCell(colNum);
              element.setCellType(element.CELL_TYPE_STRING);
              element.setCellValue(data);

              FileOutputStream fos = new FileOutputStream("C:\\Users\\target.xlsx");
              wb.write(fos);

        }
    }
