package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.Date;
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
             * Apache POI also offers some builtin formats, you can find these at :
             * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
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
            //as an example, we want to put 000x in front of the cell value
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

            

            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
