package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture08_ColumnsSizing
{
    public static void main(String[] args)
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture8.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("First Sheet");
            //Create a row, the first row will have index 0
            Row r = sheet.createRow(0);
            //Create some cells, column A is index 0, B is 1 etc
            Cell c = r.createCell(0);
            c.setCellValue("Column A");
            c = r.createCell(1);
            c.setCellValue("Column B");
            c = r.createCell(2);
            c.setCellValue("Column C");
            //Create some more rows in a loop
            for (int i = 1; i < 10; i++) {
                r = sheet.createRow(i);
                c = r.createCell(0);
                c.setCellValue( i + "A");
                c = r.createCell(1);
                c.setCellValue( i + "B");
                c = r.createCell(2);
                c.setCellValue( i + "C");
            }
            /**
             * Columns are sized in 1/256th of a character
             * We'll create a simple method in our ExcelUtils class to size columns in a nr of characters
             * The maximum column size is 255 characters
             */
            ExcelUtils.setColSize(sheet, 0, 5);
            ExcelUtils.setColSize(sheet, 1, 5);
            ExcelUtils.setColSize(sheet, 2, 5);

            /**
             * You might want to make columns narrow, but still see the column headers correctly
             * This can be done by rotating the text in the first row
             */
            ExcelUtils.angleCell(sheet, 0, 0, 45);
            ExcelUtils.angleCell(sheet, 0, 1, 45);
            ExcelUtils.angleCell(sheet, 0, 2, 45);

            /**
             * If you don't know what the size of a column should be, you can have it auto-sized
             * This will be similar to double clicking the border in excel.
             * The column will be sized to fit the widest cell it contains
             * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Sheet.html#autoSizeColumn-int-
             */
            sheet.autoSizeColumn(0);

            //Save the workbook to the filesystem
            wb.write(fileOut);
            System.out.println("Saved Excell file to  : " + filePath);
        }
        catch (IOException ex) {
            System.out.println("The file could not be written : " + ex.getMessage());
        }
    }

}
