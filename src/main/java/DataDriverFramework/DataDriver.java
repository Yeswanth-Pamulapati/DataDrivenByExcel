package DataDriverFramework;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

public class DataDriver {

    public List<HashMap<String, Object>> getData(String sheetName, String itemName) throws IOException {
        FileInputStream inputFile = new FileInputStream("C:\\Users\\yashn\\Downloads\\SaleData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputFile);
        int numberOfSheets = workbook.getNumberOfSheets();
        List<HashMap<String,Object>> hashMapList = new ArrayList<>();
        for (int i = 0; i < numberOfSheets; i++) {

            if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {

                XSSFSheet sheet = workbook.getSheetAt(i);
                Iterator<Row> rows = sheet.iterator();
                Row firstRow = rows.next();
                Iterator<Cell> cell = firstRow.cellIterator();
                int k = 0;
                int column = 0;
                while (cell.hasNext()) {
                    Cell value = cell.next();
                    if (value.getStringCellValue().equalsIgnoreCase("item")) {
                        column = k;

                    }
                    k++;
                }
                System.out.println("Column number for Item: " + column);
                while (rows.hasNext()) {
                    HashMap<String, Object> excelData = new HashMap<>();
                    Row row = rows.next();
                    if (row.getCell(column)!=null && row.getCell(column).getStringCellValue().equalsIgnoreCase(itemName)) {
                        Iterator<Cell> theCell = row.cellIterator();
                        int cellIndex = 0;
                        Cell currentCell;
                        while (theCell.hasNext()) {
                            currentCell = theCell.next();
                            String headerCell = firstRow.getCell(cellIndex).getStringCellValue().trim();
                            CellType celltype = currentCell.getCellType();
                            if(celltype == CellType.FORMULA){
                                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                                CellValue cellValue = evaluator.evaluate(currentCell);
                                //System.out.print(currentCell.getCellType() + ":" + cellValue.getNumberValue() + "\n");
                                excelData.put(headerCell,cellValue.getNumberValue());
                            }
                            else if(celltype == CellType.NUMERIC) {
                                if (DateUtil.isCellDateFormatted(currentCell)) {
                                    //System.out.print(currentCell.getCellType() + ":" + currentCell.getLocalDateTimeCellValue() + "\n");
                                    excelData.put(headerCell,currentCell.getLocalDateTimeCellValue());
                                } else {
                                    //System.out.print(currentCell.getCellType() + ":" + currentCell.getNumericCellValue() + "\n");
                                    excelData.put(headerCell,currentCell.getNumericCellValue());
                                }
                            }
                            else {
                                String currentCellValue = String.valueOf(currentCell);
                                //System.out.print(currentCell.getCellType() + ":" + currentCellValue + "\n");
                                excelData.put(headerCell,currentCellValue);
                            }

                            cellIndex++;

                        }
                        hashMapList.add(excelData);
                    }

                    }



                }
            }
        return hashMapList;
        }


        }

