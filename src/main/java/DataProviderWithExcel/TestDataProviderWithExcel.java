package DataProviderWithExcel;


import org.testng.annotations.Test;
import java.util.HashMap;

public class TestDataProviderWithExcel extends DataProviderWithExcelData {
    @Test(dataProvider="getDataFromExcel")
    public void getData(HashMap<String, Object> data){
            System.out.println(data.get("FirstName"));
            System.out.println(data.get("LastName"));
            System.out.println(data.get("EmailAddress"));

    }
}
