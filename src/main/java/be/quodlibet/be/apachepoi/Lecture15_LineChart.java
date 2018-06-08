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
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ScatterChartSeries;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.SheetBuilder;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.charts.XSSFChartLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarkerStyle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.STMarkerStyle;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNoFillProperties;


/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture15_LineChart
{
    public static void main(String[] args) throws ParseException
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture15.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("Data Sheet");
            //Create the same table as in Lecture 11 - Data Formatting
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
            List<String> headers = Arrays.asList("Start Date", "Start Time", "End Date", "End Time", "Location", "Distance", "Duration", "Average Speed");
            List<List<Object>> data = new ArrayList();
            SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss");
            String durationFormula = "=INDIRECT(\"C\"&ROW())-INDIRECT(\"B\"&ROW())";
            String speedFormula = "=INDIRECT(\"F\"&ROW())/(INDIRECT(\"G\"&ROW())*86400/3600)";
            data = DataUtils.getRandomRunningResults(25, durationFormula, speedFormula);
            HashMap<Integer, CellStyle> columnStyles = new HashMap();
            CellStyle dateStyle = sheet.getWorkbook().createCellStyle();
            dateStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("m/d/yy"));
            columnStyles.put(0, dateStyle);//Start Date
            columnStyles.put(2, dateStyle);//End Date
            CellStyle timeStyle = sheet.getWorkbook().createCellStyle();
            timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss"));
            columnStyles.put(1, timeStyle);//Start Time
            columnStyles.put(3, timeStyle);//End Time
            columnStyles.put(6, timeStyle);//Duration
            CellStyle distanceStyle = sheet.getWorkbook().createCellStyle();
            distanceStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km\""));
            columnStyles.put(5, distanceStyle);//Distance
            CellStyle speedStyle = sheet.getWorkbook().createCellStyle();
            speedStyle.setDataFormat(createHelper.createDataFormat().getFormat("##.## \"km/h\""));
            columnStyles.put(7, speedStyle);//Speed
            ExcelUtils.formatAsTable(sheet, 0, 0, "TableStyleDark2", "RunningResults", headers, data, columnStyles);

            /**
             * Line Chart
             * Create a line chart plotting the distance and speed against the dates
             */
            XSSFSheet chartSheet = wb.createSheet("Chart Sheet");
            //Create a lineChartDrawing object where the sheet will be drawn
            XSSFDrawing lineChartDrawing = chartSheet.createDrawingPatriarch();
            //Creata an lineChartAnchor where the lineChartDrawing or lineChart will be placed
            //the first 2 coordinates are the x and y coordinate inside the first cell
            //the next 2 are the coordinates inside the las cell
            //the next two are the start row and column (in our case row 1, column 1)
            //the next two are the end row and column (in our case row 25, column 20)
            XSSFClientAnchor lineChartAnchor = lineChartDrawing.createAnchor(0, 0, 0, 0, 0, 0, 25, 20);
            XSSFChart lineChart = lineChartDrawing.createChart(lineChartAnchor);
            XSSFChartLegend lineChartLegend = lineChart.getOrCreateLegend();
            lineChartLegend.setPosition(LegendPosition.BOTTOM);
            LineChartData lineChartData = lineChart.getChartDataFactory().createLineChartData();


            // Use a category axis for the bottom axis.
            ChartAxis lineChartBottomAxis = lineChart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
            lineChartBottomAxis.setNumberFormat("mm/dd/yyyy");
            ValueAxis lineChartLeftAxis = lineChart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
            lineChartLeftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            lineChartLeftAxis.setNumberFormat(distanceStyle.getDataFormatString());

            //Create a label range for the Dates
            ChartDataSource<String> labelRange = DataSources.fromStringCellRange(sheet, new CellRangeAddress(1, data.size(), 0, 0));

            //The Distance Series

            //Add a data series for the distance
            ChartDataSource<Number> distanceRange = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, data.size() + 1, 5, 5));
            lineChartData.addSeries(labelRange, distanceRange);
            //Give the series the correct title, use a reference to the column header from the data sheet
            lineChartData.getSeries().get(lineChartData.getSeries().size() - 1).setTitle(new CellReference("Data Sheet", 0, 5, true, true));



            //plot the lineChart
            lineChart.plot(lineChartData, lineChartBottomAxis, lineChartLeftAxis);
            lineChart.setTitleText("Distance by Date");

            //The default line lineChart will have a smoothed out solid line
            //We can change this using by assigning a specific marker and setting the smooth value
            CTPlotArea plotArea = lineChart.getCTChart().getPlotArea();
            //Create a CTBoolean object that is false
            CTBoolean ctFalse = CTBoolean.Factory.newInstance();
            ctFalse.setVal(false);
            //create a marker
            CTMarker ctMarker = CTMarker.Factory.newInstance();
            //create a marker style with a start symbols
            CTMarkerStyle starMarkerStyle = CTMarkerStyle.Factory.newInstance();
            starMarkerStyle.setVal(STMarkerStyle.STAR);
            //Assign the symbol to the marker
            ctMarker.setSymbol(starMarkerStyle);

            //now, set the markerstyle and unset the smooth line setting for the first series of the graph
            plotArea.getLineChartArray()[0].getSerArray()[0].setSmooth(ctFalse);
            plotArea.getLineChartArray()[0].getSerArray()[0].setMarker(ctMarker);

            /**
             * Scatter Chart
             * Create a line chart plotting the XX and XX against the XX
             */
            XSSFDrawing scatterChartDrawing = chartSheet.createDrawingPatriarch();
            XSSFClientAnchor scatterChartAnchor = scatterChartDrawing.createAnchor(0, 0, 0, 0, 0, 21, 25, 41);
            XSSFChart scatterChart = scatterChartDrawing.createChart(scatterChartAnchor);
            //XSSFChartLegend scatterChartLegend = scatterChart.getOrCreateLegend();
            //scatterChartLegend.setPosition(LegendPosition.BOTTOM);
            ScatterChartData scatterChartData = scatterChart.getChartDataFactory().createScatterChartData();

            // Use a category axis for the bottom axis.
            ChartAxis scatterChartBottomAxis = scatterChart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
            lineChartBottomAxis.setNumberFormat("mm/dd/yyyy");
            ValueAxis scatterChartLeftAxis = scatterChart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
            scatterChartLeftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            scatterChartLeftAxis.setNumberFormat(distanceStyle.getDataFormatString());
            //We add the same series to the scatter chart as we added to the line chart
            //this method is called addSerie (no s at the end)
            scatterChartData.addSerie(labelRange, distanceRange);
            scatterChart.plot(scatterChartData, scatterChartBottomAxis, scatterChartLeftAxis);

            //now, by default the scatter chart will use smooth lines, if you want a scatter chart without lines
            //use below code
            CTPlotArea scatterPlotArea = scatterChart.getCTChart().getPlotArea();

            //set the no fill property for the line so the line is not visible
            CTNoFillProperties noFill = CTNoFillProperties.Factory.newInstance();
            noFill.setNil();
            scatterPlotArea.getScatterChartArray()[0].getSerArray()[0]
                    .addNewSpPr()
                    .addNewLn()
                    .setNoFill(noFill);
            //Use the same marker as the line chart did
            scatterPlotArea.getScatterChartArray()[0].getSerArray()[0].setMarker(ctMarker);
            //Hide the different types of labels
            CTDLbls labels = scatterPlotArea.getScatterChartArray()[0].getSerArray()[0].addNewDLbls();
            labels.addNewShowSerName().setVal(false);//Don't show series name
            labels.addNewShowVal().setVal(false);//don't show X value
            labels.addNewShowCatName().setVal(false);//don't show Y value
            


            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
            testOneSeriePlot();
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }

    }

    public static void testOneSeriePlot()
    {
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";

        String filePath = excellFolder + "scatterexample.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
        Object[][] plotData = {
            {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"},
            {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}};
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = new SheetBuilder(wb, plotData).build();
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing
                .createAnchor(0, 0, 0, 0, 1, 1, 10, 30);
        Chart chart = drawing.createChart(anchor);

        ChartAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(
                AxisPosition.BOTTOM);
        ChartAxis leftAxis = chart.getChartAxisFactory().createValueAxis(
                    AxisPosition.LEFT);


        ScatterChartData scatterChartData = chart.getChartDataFactory()
                .createScatterChartData();

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet,
                                                                     CellRangeAddress.valueOf("A1:J1"));
        ChartDataSource<Number> ys = DataSources.fromNumericCellRange(
                sheet, CellRangeAddress.valueOf("A2:J2"));
        ScatterChartSeries series = scatterChartData.addSerie(xs, ys);

        /*assertNotNull(series);
         assertEquals(1, scatterChartData.getSeries().size());
        assertTrue(scatterChartData.getSeries().contains(series));
                */
            chart.plot(scatterChartData, bottomAxis, leftAxis);
            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
