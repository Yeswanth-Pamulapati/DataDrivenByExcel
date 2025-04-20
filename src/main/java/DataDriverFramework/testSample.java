package DataDriverFramework;

import java.io.IOException;

public class testSample {
    public static void main(String[] args) throws IOException {
        DataDriver dataDriver = new DataDriver();
        System.out.println(dataDriver.getData("Sales-2022","Television"));
    }
}
