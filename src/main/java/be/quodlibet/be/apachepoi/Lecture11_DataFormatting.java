package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture11_DataFormatting
{
    public static void main(String[] args) throws ParseException
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture11.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("First Sheet");

            /**
             * Data formats can be assigned to cells just as they are in excell
             * You can find the notation by using the Custom Number format in the excel format dialog
             * Apache POI also offers some builtin formats, you can find these at : https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
             */
            //Create a cell style with a predefined date format
            //assign a built in date format to this cell
            CellStyle dateStyle = sheet.getWorkbook().createCellStyle();
            dateStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy"));

            //Create a custom data format, the CreateHelper class helps creating various things, among these things is a dataformat
            CellStyle pctStyle = sheet.getWorkbook().createCellStyle();
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
            //We'll create a format for a percentage with 2 digit precision
            pctStyle.setDataFormat(createHelper.createDataFormat().getFormat("0.00%"));

            //Create a custom format that displays the numeric value with a specific prefix,
            //as an example, we wan't to put 000x in front of the cell value
            // 11111 would be displayed as 000x11111
            // 12345 as 000x12345
            CellStyle zeroxStyle = sheet.getWorkbook().createCellStyle();
            zeroxStyle.setDataFormat(createHelper.createDataFormat().getFormat("000x######"));

            //Create a row, the first row will have index 0
            Row r = sheet.createRow(0);
            //Create a cell that contains a date
            Cell c = r.createCell(0);
            c.setCellValue(new Date());
            //Assign the date style
            c.setCellStyle(dateStyle);
            //Create a cell that contains a percentage
            c = r.createCell(1);
            c.setCellValue(0.2525);//25.25%
            c.setCellStyle(pctStyle);
            //Create a cell that contains a value formatted as 000x######
            c = r.createCell(2);
            c.setCellValue(99999);//000x99999
            c.setCellStyle(zeroxStyle);

            /**
             * Now we'll create a method similar to the formatAsStringTable from previous lecture,
             * but we will add the ability to add a specific format to each column, and allow different
             * types of objects instead of just strings
             */
            //Create a List of column headers
            List<String> headers = Arrays.asList("Start Date", "Start Time", "End Date", "End Time", "Location", "Distance", "Duration", "Average Speed");
            /**
             * The last two columns, duration and speed, we'll let excell calculate using a formula
             */
            List<List<Object>> data = new ArrayList();
            //Create a formatter to easily add dates and times
            SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss");

            /**
             * For the duration column, we'll use a formula to calculate the duration between start and finish
             * Since we don't really know yet what row this will end up in in the spreadsheet, we'll use a formula that references the
             * current row. That can be done using the INDIRECT formula, this formula allows you to create the string that makes up a formula
             * during the evaluation
             * Column B will contain the Start Time, and Column D will contain the end time
             * This will only work if we start the table in column A (or 0) otherwise we need to adapt the formula
             * We'll use this formula
             * =INDIRECT("D"&ROW())-INDIRECT("B"&ROW())
             * The evaluated formula for row 1 will be
             * =D1-B1
             */
            String durationFormula = "=INDIRECT(\"C\"&ROW())-INDIRECT(\"B\"&ROW())";
            /**
             * The speed will be calculated as km/h
             * Excell times are always in 86400th of a second
             * 1 second = 1/86400 = 0.0000115740740740741
             * 1 minute = 60/86400 = 0.000694444444444444
             * 1 hour = 3600/86400 = 0.0416666666666667
             * 1 day = 86400/86400 = 1
             * So to calculate the speed, we need the following formula
             * we want km/h, so we first convert excell time to seconds G1 * 86400
             * then convert to hours G1 * 86400 * 3600
             * Formula is
             * =F1/(G1*86400/3600)
             * Use INDIRECT, because we don't know what the rows will be
             * =INDIRECT("F"&ROW())/(INDIRECT("G"&ROW())*86400/3600)
             */
            String speedFormula = "=INDIRECT(\"F\"&ROW())/(INDIRECT(\"G\"&ROW())*86400/3600)";

            //Create some data records representing a short run in a city in belgium
            data = DataUtils.getRandomRunningResults(25, durationFormula, speedFormula);

            /**
             * In our ExcelUtils class, we will create a method that can take the same parameters as formatAsStringTable,
             * but we'll allow mixed objects to be passed as values, and we'll also allow to pass column formatting to the columns.
             */

            //Create a map of styles, they are mapped to the columns by their index
            HashMap<Integer, CellStyle> columnStyles = new HashMap();
            //Style for date, reuse the style we used before
            columnStyles.put(0, dateStyle);//Start Date
            columnStyles.put(2, dateStyle);//End Date
            //Time Style
            CellStyle timeStyle = sheet.getWorkbook().createCellStyle();
            timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss"));
            columnStyles.put(1, timeStyle);//Start Time
            columnStyles.put(3, timeStyle);//End Time
            columnStyles.put(6, timeStyle);//Duration

            //Distance Style
            CellStyle distanceStyle = sheet.getWorkbook().createCellStyle();
            distanceStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km\""));
            columnStyles.put(5, distanceStyle);//Distance

            //Speed Style
            CellStyle speedStyle = sheet.getWorkbook().createCellStyle();
            speedStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km/h\""));
            columnStyles.put(7, speedStyle);//Speed

            //Create the table
            ExcelUtils.formatAsTable(sheet, 5, 0, "TableStyleDark2", "RunningResults", headers, data, columnStyles);


            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
