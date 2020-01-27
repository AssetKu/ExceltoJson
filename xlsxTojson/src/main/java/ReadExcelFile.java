import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcelFile {
    private static XSSFWorkbook mybook;
    static String fileLocation = "C:\\Users\\u10074\\Desktop\\Новая папка\\book1.xlsx";

    public static void main(String[] args) {
        try {
            File newFile = new File(fileLocation);
            FileInputStream fIO = new FileInputStream(newFile);
            JSONArray array = new JSONArray();
            mybook = new XSSFWorkbook(fIO);
            XSSFSheet mySheet = mybook.getSheetAt(0);

            DataFormatter dataFormatter = new DataFormatter();

            int numOfRows = mySheet.getPhysicalNumberOfRows();

            for (int i = 0 ; i < numOfRows; i++) {
                XSSFRow row = mySheet.getRow(i);
                JSONObject myObject = new JSONObject();

                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);
                    String cellValue = dataFormatter.formatCellValue(cell);
                    if (j == 0) {

                        myObject.put("Country_Code", cellValue.substring(0,2));
                        myObject.put("CountryName", cellValue.substring(3));
                        //System.out.println(myObject);
                    } else if (j == 1) {

                        myObject.put("City_Type", cellValue.substring(0, cellValue.indexOf(" ")));
                        myObject.put("CityName", cellValue.substring(cellValue.indexOf(" ") + 1 ));
                       // System.out.println(myObject);
                    }
                }
                array.put(myObject);
            }
            System.out.println(array);

            fIO.close();
        } catch (FileNotFoundException ef) {
            ef.printStackTrace();
        } catch (IOException ei) {
            ei.printStackTrace();
        }
    }
}