package Tiwari.Cucumber;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelSheet {

	static XSSFSheet sheet= null;
	final static String path ="C:\\Users\\User\\cucucmberWorkshop\\Excel\\book.xlsx";
	public static int rowNum =0;
	 public static int  col_Num = 0;

  
	public static void main(String[] args) throws IOException   
	{  
	  
	String vOutput=ReadCellData("Data","TC_002","Language");   
	System.out.println(vOutput);  
	}  
	//method defined for reading a cell  
	public static String ReadCellData(String sheetName, String rowname, String cellName) throws IOException  
	{ 
		
	
	String value=null;          //variable for storing the cell value  
	Workbook wb=null;           //initialize Workbook null  
	
	try {
	FileInputStream fis=new FileInputStream("C:\\Users\\User\\eclipse-workspace_new\\Cucumber\\book1.xlsx");   
	wb=new XSSFWorkbook(fis); 
	}
	catch (FileNotFoundException e) {
		e.printStackTrace();
	}
     sheet = (XSSFSheet) wb.getSheet(sheetName);
     Row row = sheet.getRow(0);
     for(int m=1 ; m<= sheet.getPhysicalNumberOfRows() ;m++) {
     	
 	 	 Row rows = CellUtil.getRow(m, sheet);
 	     Cell cell = CellUtil.getCell(rows,0);
 	 
 	    //ROW 
 	     if(cell.toString().equalsIgnoreCase(rowname)) {
 	    	 rowNum = m;
 	    	 break;
 	     }    
    }
     //Column
     for (int i = 0; i <= row.getLastCellNum(); i++) {
         if (row.getCell(i).getStringCellValue().trim().equals(cellName))
         {
             col_Num = i;
             break;
         }
        
     } 
    
     	Row row1=sheet.getRow(rowNum); 
		Cell cell = row1.getCell(col_Num); //getting the cell representing the given column  
		value=cell.getStringCellValue();    //getting cell value  
		return value;               //returns the cell value  */
	} 
	
		
	}
	

