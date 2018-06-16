package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture13_ConditionalFormatting
{
    public static void main(String[] args) throws ParseException
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture12.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("First Sheet");
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();

            /**
             * We reuse the table of the previous lecture that contains the Running results
             */
            List<String> headers = Arrays.asList("Start Date", "Start Time", "End Date", "End Time", "Location", "Distance", "Duration", "Average Speed");
            List<List<Object>> data = new ArrayList();
            SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss");
            String durationFormula = "=INDIRECT(\"C\"&ROW())-INDIRECT(\"B\"&ROW())";
            String speedFormula = "=INDIRECT(\"F\"&ROW())/(INDIRECT(\"G\"&ROW())*86400/3600)";

            //Create some data records representing a short run in a city in belgium
            data = DataUtils.getRandomRunningResults(25, durationFormula, speedFormula);



            HashMap<Integer, CellStyle> columnStyles = new HashMap();
            CellStyle dateStyle = sheet.getWorkbook().createCellStyle();
            CellStyle timeStyle = sheet.getWorkbook().createCellStyle();
            CellStyle distanceStyle = sheet.getWorkbook().createCellStyle();
            CellStyle speedStyle = sheet.getWorkbook().createCellStyle();
            dateStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy"));
            timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss"));
            distanceStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km\""));
            speedStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km/h\""));

            columnStyles.put(0, dateStyle);//Start Date
            columnStyles.put(1, timeStyle);//Start Time
            columnStyles.put(2, dateStyle);//End Date
            columnStyles.put(3, timeStyle);//End Time
            columnStyles.put(5, distanceStyle);//Distance
            columnStyles.put(6, timeStyle);//Duration
            columnStyles.put(7, speedStyle);//Speed
            ExcelUtils.formatAsTable(sheet, 0, 0, "TableStyleDark2", "RunningResults", headers, data, columnStyles);

            /**
             * We assign conditional formatting Color Scale rule to the Speed column
             * For this we create another method in the ExcelUtils class called colorScaleRangeNumber
             * Red for the low value #ff0000
             * Orange for the Median #ff9900
             * Green for high values #006600
             */
            ExcelUtils.colorScaleRangeNumber(sheet, 1, 7, 1 + data.size(), 7, "ff0000", "ff9900", "006600");
            /**
             * We assign conditional formatting Icon Set rule to the Distance column
             * There are several icon sets available, all of them will be provided in the course resources
             * Some have 4 and some have 5 icons, the method IconSetRangeNumber works best with 5
             */
            ExcelUtils.iconSetRange(sheet, 1, 5, 1 + data.size(), 5, IconMultiStateFormatting.IconSet.RATINGS_5);
            /**
             * We assign conditional formatting Data Bar rule to the Duration column
             * The color of the bar is passed as a parameter
             */
            ExcelUtils.dataBarRange(sheet, 1, 6, 1 + data.size(), 6, "006600");
            //Auto sizer all columns
            int colindex = 0;
            for (String header : headers) {

                sheet.autoSizeColumn(colindex++);
            }
            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
