package delta.exceltool;
 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
/*
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
*/

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Create a new sheet.
 * 
 * @author 
 *
 */
public class ExcelFileUpdateHSS {
 
 
    public static void main(String[] args) {
        String excelFile = "C:\\Project\\Excel\\JavaBooks5.xlsx";
         boolean bool = false; 
         File f = null;
         int rowval = 4;
         int colval = 3 ;
         
        try {
             
			f = new File(excelFile) ;
			bool = f.createNewFile() ;
			
        if (bool)
        {
            System.out.println("File is created " + excelFile);
        }
        else 
        {
        	  System.out.println("File already exists " + excelFile) ;
        }
        
        FileOutputStream fos2 = new FileOutputStream (excelFile);
        HSSFWorkbook workbook = new HSSFWorkbook();
     //    Workbook workbook = WorkbookFactory.create(fis2) ; 
           HSSFSheet sheet = workbook.createSheet("Test1");
          Row[] row = new Row[rowval]  ;
          Cell[][] cell = new  Cell[rowval][colval]  ;
           
           //create cells
           for (int i=1; i< rowval ; i++)
           {
        	   row[i]= sheet.createRow(i) ; 
        	  for (int j=0; j<colval; j++)
        	  {   	 
                 cell[i][j] = row[i].createCell(j); 
        	  }
           } 
           for (int i=1; i< rowval ; i++)
           {
                 if (i==0)
                 {
                	 //set column
                	cell[i][0].setCellValue("Id");
                	cell[i][1].setCellValue("Name");
                	cell[i][2].setCellValue("Age");
                	
                 }
                 if (i==1)
                 {
                	 //set column
                		cell[i][0].setCellValue("1");
                    	cell[i][1].setCellValue("R");
                    	cell[i][2].setCellValue("12");
                	
                 }
                 if (i== 2)
                 {
                	 //set column
                		cell[i][0].setCellValue("2");
                    	cell[i][1].setCellValue("S");
                    	cell[i][2].setCellValue("10");
                 }
                 if (i== 3)
                 {
                	 //set column
                		cell[i][0].setCellValue("3");
                    	cell[i][1].setCellValue("N");
                    	cell[i][2].setCellValue("8");
                 }
       	   }
      workbook.write(fos2);
      fos2.flush();
      fos2.close();
           
           for (int i=1; i< rowval ; i++)
           {
        	  for (int j=0; j<colval; j++)
        	  {   	 
            System.out.println("cell " + i + j +"= " + cell[i][j]); 
        	  }
           } 
        	 
        }  
        catch (Exception e)
        {
        	e.printStackTrace();
        }
           
    } 
}
           
 
 