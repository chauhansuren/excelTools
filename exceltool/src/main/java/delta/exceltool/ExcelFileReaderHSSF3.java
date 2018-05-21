package delta.exceltool;
 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Create a new sheet.
 * 
 * @author 
 *
 */
public class ExcelFileReaderHSSF3 {
 
 
    public static void main(String[] args) {
        String excelFile = "C:\\Project\\Excel\\JavaBooks5.xls";
         boolean bool = false; 
         File f = null;
         int rowval = 4;
         int colval = 3 ;
         
        try {
        
        FileInputStream fis2 = new FileInputStream (excelFile);
       HSSFWorkbook workbook = new HSSFWorkbook(fis2);
    //   Workbook workbook = WorkbookFactory.create(fis2) ; 
       //    XSSFSheet firstSheet = workbook.getSheet("Test1");
     HSSFSheet sheet = workbook.getSheetAt(0);
           
           //Sheet datatypeSheet = workbook.getSheetAt(0);
           Iterator<Row> iterator = sheet.iterator();
           
      //     Iterator<Row> iterator = firstSheet.iterator();
           
           while (iterator.hasNext()) {
               Row nextRow = iterator.next();
               Iterator<Cell> cellIterator = nextRow.cellIterator();
                
               while (cellIterator.hasNext()) {
                   Cell cell = cellIterator.next();
                    
                   switch (cell.getCellType()) {
                       case Cell.CELL_TYPE_STRING:
                           System.out.print(cell.getStringCellValue());
                           break;
                       case Cell.CELL_TYPE_BOOLEAN:
                           System.out.print(cell.getBooleanCellValue());
                           break;
                       case Cell.CELL_TYPE_NUMERIC:
                           System.out.print(cell.getNumericCellValue());
                           break;
                   }
                   System.out.print(" - ");
               }
               System.out.println();
           }
            
       //workbook.close();
           fis2.close();
       }
           
        catch (Exception e)
        {
        	e.printStackTrace();
        }
           
    } 
}
           
 
 