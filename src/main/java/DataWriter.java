
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.skyscreamer.jsonassert.JSONCompareResult;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;


public class DataWriter {

    public DataWriter() {
    }

    public static void writeData(XSSFSheet comparsionSheet, JSONCompareResult result, String id, String test_case) {
    }

    public static void writeData(XSSFSheet resultSheet, String aFalse, String id, String test_case, int i) {
    }

    public static void writeData(XSSFSheet outputSheet, String s, String ID, String test_case) {
    }

    public static void writeData(XSSFSheet resultSheet, double totalcase, double failedcase, String startTime, String endTime) {
    }


    public static void writeData(XSSFSheet comparsionSheet, String s, String s1, String id, String test_case) {
    }

    //写入Xlsx
    public static void writeXlsx(String fileName,Map<Integer,List<String[]>> map) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            for(int sheetnum=0;sheetnum<map.size();sheetnum++){
                XSSFSheet sheet = wb.createSheet(""+sheetnum);
                List<String[]> list = map.get(sheetnum);
                for(int i=0;i<list.size();i++){
                    XSSFRow row = sheet.createRow(i);
                    String[] str = list.get(i);
                    for(int j=0;j<str.length;j++){
                        XSSFCell cell = row.createCell(j);
                        cell.setCellValue(str[j]);
                    }
                }
            }
            FileOutputStream outputStream = new FileOutputStream(fileName);
            wb.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}