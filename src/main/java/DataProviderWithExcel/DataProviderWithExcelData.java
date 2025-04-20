package DataProviderWithExcel;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class DataProviderWithExcelData {

    @DataProvider
    public Object[][] getDataFromExcel() throws IOException {
        FileInputStream fis = new FileInputStream("C:\\Users\\yashn\\Downloads\\random_users.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int numerofRows = sheet.getPhysicalNumberOfRows();
        XSSFRow firstRow = sheet.getRow(0);
        int numberofColumns = firstRow.getLastCellNum();
        Object[][] data = new Object[numerofRows-1][1];
        DataFormatter dataFormatter = new DataFormatter();
        for(int i =0;i<numerofRows-1;i++){
            HashMap<String,Object> hashMap = new HashMap<>();
            XSSFRow row = sheet.getRow(i+1);
            for(int j=0;j<numberofColumns;j++){
                String header = firstRow.getCell(j).getStringCellValue();
                String value = dataFormatter.formatCellValue(row.getCell(j));
                // this formatCellValue converts the data of any type into string.
                hashMap.put(header,value);

            }
            data[i][0] = hashMap;

        }
        return data;
    }
}
