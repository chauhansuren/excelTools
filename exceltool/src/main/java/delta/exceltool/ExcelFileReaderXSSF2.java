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

import jdk.nashorn.internal.ir.RuntimeNode.Request;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Create a new sheet.
 * 
 * @author 
 *
 */
public class ExcelFileReaderXSSF2 {

    public static void main(String[] args) {
        String excelFile = "C:\\Project\\Excel\\JavaBooks6.xlsx";
         boolean bool = false; 
         File f = null;
         int rowval = 4;
         int colval = 3 ;
         
        try {
        
     //   FileInputStream fis2 = new FileInputStream (excelFile);
        FileInputStream file1 = new FileInputStream(new File(excelFile));
        System.out.println(file1);
        file1.close();    
        //Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook111 = new XSSFWorkbook(file1);

        //Get first/desired sheet from the workbook
        XSSFSheet sheet111 = workbook111.getSheetAt(0);


        Object[][] bookData_read = new String[sheet111.getLastRowNum()+1][3];

        int row_count11 = 0;

        Iterator<Row> rowIterator111 = sheet111.iterator();

        while (rowIterator111.hasNext()) 
        {
            Row row111 = rowIterator111.next();
            if (row_count11 == 0) {
                row_count11++;
            continue;

            }

            if (row_count11 > sheet111.getLastRowNum())
                break;



            Cell cell11 = row111.getCell(4);
            String cellvalue1 = "";

            if (cell11 != null && ! "".equals(cell11.getStringCellValue()) ||cell11 != null && ! "".equals(cell11.getNumericCellValue())) {

                switch (cell11.getCellType()) {

                    case Cell.CELL_TYPE_STRING:
                        //System.out.print(cell.getStringCellValue());
                        cellvalue1 = cell11.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        //System.out.print(cell.getBooleanCellValue());
                        cellvalue1 = "" + cell11.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //System.out.print(cell.getNumericCellValue());
                        cellvalue1 = "" + cell11.getNumericCellValue();
                        break;

                }}


                String module1 = cellvalue1;
                //System.out.println("module"+module1);
                //System.out.println(cellvalue);
                bookData_read[row_count11][0] = cellvalue1;


                if(row111.getCell(5)!=null)
                {
                      Cell cell2 = row111.getCell(5);
                      cellvalue1 = "";
                    if (cell2 != null && ! "".equals(cell2.getStringCellValue())) {



                switch (cell2.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        //System.out.print(cell.getStringCellValue());
                        cellvalue1 = cell2.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        //System.out.print(cell.getBooleanCellValue());
                        cellvalue1 = "" + cell2.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //System.out.print(cell.getNumericCellValue());
                        cellvalue1 = "" + cell2.getNumericCellValue();
                        break;
                }  }}
                String submodule = cellvalue1;
                //System.out.println("submodule"+submodule);

               // System.out.println(cellvalue);
                bookData_read[row_count11][1] = cellvalue1;



               if(row111.getCell(5)!=null && ! "".equals(row111.getCell(5).getStringCellValue()) )
               {
//               if(row.getCell(7)!=null && ! "".equals(row.getCell(7).getStringCellValue()))
//              {
//                    
                Cell  cell3 = row111.getCell(8);
                cellvalue1 = "";
                if (cell3 != null && ! "".equals(cell3.getStringCellValue())) {
                switch (cell3.getCellType()) {
                    case Cell.CELL_TYPE_STRING:

                        cellvalue1 = cell3.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:

                        cellvalue1 = "" + cell3.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:

                        cellvalue1 = "" + cell3.getNumericCellValue();
                        break;}}


               String  temp = cellvalue1;
               bookData_read[row_count11][2] = cellvalue1;

               }
                else{

                   continue;
                }


               row_count11++;



        }
  //      workbook111.close();

request.setAttribute("modulesList", bookData_read);
RequestDispatcher rd = request.getRequestDispatcher("/All_Modules.jsp");
    rd.forward(request, response);  
 
 