package UploadDownloadExcelFiles;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ParseExcel {


	public Object getData(String sheetName, String fruitName) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\yashn\\Downloads\\download.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		//XSSFSheet workSheet = workbook.getSheetAt(0);
		int numberOfSheets = workbook.getNumberOfSheets();
		int NumberOfColumns;
        int priceColumnNumber = 0,fruitColumnNumber = 0,fruitRowNumber=0,priceRowNumber=0;
        int price = 0;
		for (int i = 0; i < numberOfSheets; i++) {

            if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {

                XSSFSheet sheet = workbook.getSheetAt(i);
                Iterator<Row> rows = sheet.iterator();
                Row firstRow = rows.next();
                NumberOfColumns = firstRow.getLastCellNum();
		for(int j=0; j< NumberOfColumns; j++) {
			if(firstRow.getCell(j).getStringCellValue().equalsIgnoreCase("price")) {
				priceColumnNumber = j;
				
			}
			else if(firstRow.getCell(j).getStringCellValue().equalsIgnoreCase("fruit_name")) {
				fruitColumnNumber = j;
			}
				
			}
		while(rows.hasNext()){
			
			Row currentRow = rows.next();
			if(currentRow.getCell(fruitColumnNumber).getStringCellValue().equalsIgnoreCase(fruitName)) {
				DataFormatter dataFormatter = new DataFormatter();
				String value = dataFormatter.formatCellValue(currentRow.getCell(priceColumnNumber));
				price = Integer.parseInt(value);
				
			};
		}
			}
		
		}
		return price;	 
		
	}
	
	public Boolean updateSheet(String sheetName, int row, int column, String valueToUpdate) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\yashn\\Downloads\\download.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		XSSFRow CurrentRow = sheet.getRow(row);
		Cell getCell = CurrentRow.getCell(column);
		getCell.setCellValue(valueToUpdate);
		FileOutputStream fos = new FileOutputStream("C:\\Users\\yashn\\Downloads\\download.xlsx");
		workbook.write(fos);
		fos.close();
		return true;
		
	}
	
	public int getRow(String sheetName, String fruitName) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\yashn\\Downloads\\download.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		Iterator<Row> rows = sheet.iterator();
		Row returnRow = null;
		while(rows.hasNext()) {
			Row currentRow = rows.next();
			Iterator<Cell> cells = currentRow.cellIterator();
			while(cells.hasNext()) {
				Cell currentCell = cells.next();
				DataFormatter dataFormatter = new DataFormatter();
				String value = dataFormatter.formatCellValue(currentCell);
				if(value.equalsIgnoreCase(fruitName)) {
					returnRow = currentRow;
				}
			}
		}
		return returnRow.getRowNum();	
	}
	public int getColumnNumber(String sheetName,String columnName) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\yashn\\Downloads\\download.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		int columnCounter=0,columnNumber=0;
		Iterator<Row> rows = sheet.iterator();
		Row firstRow = rows.next();
		Iterator<Cell> currentRowCells = firstRow.cellIterator();
		while(currentRowCells.hasNext()) {
			Cell currentCell = currentRowCells.next();
			DataFormatter dataFormatter = new DataFormatter();
			String value = dataFormatter.formatCellValue(currentCell);
			if(value.equalsIgnoreCase(columnName)) {
				columnNumber = columnCounter;
			}
			 columnCounter++;
			}
		return columnNumber;
		}
}
