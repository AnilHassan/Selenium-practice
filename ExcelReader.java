package session_2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jopendocument.dom.spreadsheet.ColumnStyle;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.SpreadSheet;


public class ExcelReader {
	
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException{


	//open Excel

	 Workbook wb=WorkbookFactory.create(new FileInputStream("./Data/Daily Tasks updates.xlsx"));
	Sheet ss = wb.getSheet("January 2018");
	 
	 for(int row=0;row < ss.getLastRowNum();row++){
		 
		 for(int coloumn=1;coloumn < ss.getRow(row).getLastCellNum();coloumn++){
			CellType type = ss.getRow(row).getCell(coloumn).getCellTypeEnum();
			if (type == CellType.STRING){
				String value = ss.getRow(row).getCell(coloumn).toString();
				System.out.println(value);
			}
			
			else if(type == CellType.NUMERIC ||type == CellType.FORMULA){
				int value=(int) ss.getRow(row).getCell(coloumn).getNumericCellValue();
				
				System.out.println(value);
				
			}
			
			else if(type == CellType.BLANK){
				System.out.println("Blank Cell");
			}
		 }
	 }
	  
	 
	 
	  
//	  Workbook wb1=WorkbookFactory.create(new FileInputStream("./Data/Input1.xls"));
//	  wb1.getSheet("Sheet1").createRow(0).createCell(0).setCellValue("Bye");
//	
//	  wb1.write(new FileOutputStream("./Data/Input1.xls"));
//	  
//	  File file = new File("./Data/Input2.ods");

	  
//org.jopendocument.dom.spreadsheet.Sheet sh=SpreadSheet.createFromFile(file).getSheet(0);
//int ColCount = sh.getColumnCount();
//int RowCount = sh.getRowCount();
//MutableCell cell = null;
//for(int RowNum = 0; RowNum < sh.getRowCount(); RowNum++)
//{
//  //Iterating through each column
// 
//  for( int ColNum = 0 ;ColNum <sh.getColumnCount(); ColNum++)
//  {
// cell=sh.getCellAt(ColNum, RowNum);
//	Object val = cell.getValue();
//	System.out.println(val);
//	cell.setValue(val);
//
//  }}
	
	}
}