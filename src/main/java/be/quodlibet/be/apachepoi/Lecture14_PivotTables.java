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
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture13_PivotTables
{
    public static void main(String[] args) throws ParseException
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture13.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("First Sheet");
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();

            /**
             * We reuse the table of the previous lecture that contains the Running results
             * We add an additional calculated field called "Month" so we can easily group by month
             */
            List<String> headers = Arrays.asList("Start Date", "Start Time", "End Date", "End Time", "Location", "Distance", "Duration", "Average Speed", "Month");
            List<List<Object>> data = new ArrayList();
            SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss");
            String durationFormula = "=INDIRECT(\"C\"&ROW())-INDIRECT(\"B\"&ROW())";
            String speedFormula = "=INDIRECT(\"F\"&ROW())/(INDIRECT(\"G\"&ROW())*86400/3600)";
            String monthFormula = "=TEXT(INDIRECT(\"A\"&ROW()),\"mmmm\")";

            //Create some data records representing a short run in a city in belgium
            data = DataUtils.getRandomRunningResults(100, durationFormula, speedFormula, monthFormula);
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
            ExcelUtils.colorScaleRangeNumber(sheet, 1, 7, 1 + data.size(), 7, "ff0000", "ff9900", "006600");
            ExcelUtils.iconSetRange(sheet, 1, 5, 1 + data.size(), 5, IconMultiStateFormatting.IconSet.RATINGS_5);
            ExcelUtils.dataBarRange(sheet, 1, 6, 1 + data.size(), 6, "006600");
            //Auto size all columns
            int colindex = 0;
            for (String header : headers) {
                sheet.autoSizeColumn(colindex++);
            }
            /**
             * First Pivot Table
             * Figure out how much we ran in each Location
             */
            //We'll put the pivot table in a new sheet
            XSSFSheet pivotSheet = wb.createSheet("Location Pivot Table");
            //The range we'll create a pivot from is colum A to H,
            //in rows 1 to 101 (100 data records and 1 header)
            AreaReference dataArea = new AreaReference(
                    new CellReference( 0, 0),//A1
                    new CellReference(data.size(), 8)//H101
            );
            //The location where we'll place the pivot table, first row and column in the Pivot Sheet
            //Make sure to always start a pivot table at least at row 3 , excell can't handle pivot tables that start in rows 0,1,2
            CellReference pivotTargetCell = new CellReference(3, 0);
            //Create the pivot table (pass sheet as the third parameter  because that's what the dataArea refers to)
            XSSFPivotTable pivotTable = pivotSheet.createPivotTable(dataArea, pivotTargetCell, sheet);
            //Set the label column to the Location
            pivotTable.addRowLabel(4);//Location
            //Set the label for that row
            pivotTable.getCTPivotTableDefinition().setRowHeaderCaption("Location");
            //Setup the Value columns
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 5, "Distance"); //Total Distance
            pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 7, "Speed"); //Average Speed
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 6, "Duration"); //Total Duration

            //Apply the correct data format to the columns
            //We can't use the normal cells to assign the format, so we reference the CTPivotTable and assign the formats to the fields
            CTDataFields pivotFields = pivotTable.getCTPivotTableDefinition().getDataFields();
            for (CTDataField field : pivotFields.getDataFieldList()) {
                switch (field.getName()) {
                    case "Distance":
                        field.setNumFmtId(distanceStyle.getDataFormat());
                        break;
                    case "Speed":
                        field.setNumFmtId(speedStyle.getDataFormat());
                        break;
                    case "Duration":
                        field.setNumFmtId(timeStyle.getDataFormat());
                        break;
                }
            }

            /**
             * Second Pivot Table
             * Figure out how much we ran in each Month, and add a filter by location
             */
            XSSFSheet pivotSheet2 = wb.createSheet("Month Pivot Table");
            dataArea = new AreaReference(
                    new CellReference(0, 0),//A1
                    new CellReference(data.size(), 8)//H101
            );
            pivotTargetCell = new CellReference(3, 0);
            pivotTable = pivotSheet2.createPivotTable(dataArea, pivotTargetCell, sheet);
            //Set the label column to the Month
            pivotTable.addRowLabel(8);//Month
            //Set the label for that row
            pivotTable.getCTPivotTableDefinition().setRowHeaderCaption("Month");
            //Setup the Value columns
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 5, "Distance"); //Total Distance
            pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 7, "Speed"); //Average Speed
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 6, "Duration"); //Total Duration

            //Apply the correct data format to the columns
            //We can't use the normal cells to assign the format, so we reference the CTPivotTable and assign the formats to the fields
            pivotFields = pivotTable.getCTPivotTableDefinition().getDataFields();
            for (CTDataField field : pivotFields.getDataFieldList()) {
                switch (field.getName()) {
                    case "Distance":
                        field.setNumFmtId(distanceStyle.getDataFormat());
                        break;
                    case "Speed":
                        field.setNumFmtId(speedStyle.getDataFormat());
                        break;
                    case "Duration":
                        field.setNumFmtId(timeStyle.getDataFormat());
                        break;
                }
            }
            //Add a filter on the Location
            pivotTable.addReportFilter(4);




            //Save the workbook to the filesystem
            wb.write(fileOut);

            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
