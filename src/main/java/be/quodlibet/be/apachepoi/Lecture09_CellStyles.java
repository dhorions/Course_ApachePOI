package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture09_CellStyles
{
    public static void main(String[] args)
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture9.xlsx";
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
                c.setCellValue(i + "A");
                c = r.createCell(1);
                c.setCellValue(i + "B");
                c = r.createCell(2);
                c.setCellValue(i + "C");
            }
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            /**
             * Create Styles to apply to cells. Cell Styles can be created by the workbook.
             * To see all methods for a cell style, check the documentation
             * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/CellStyle.html
             */
            //Create a cell style for the first row of our spreadsheet
            XSSFCellStyle headerCellStyle = wb.createCellStyle();
            /**
             * Alignment
             */
            headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
            headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            /**
             * Borders - left,top,right,bottom
             */
            //Border Style

            headerCellStyle.setBorderLeft(BorderStyle.THICK);
            headerCellStyle.setBorderTop(BorderStyle.THICK);
            headerCellStyle.setBorderRight(BorderStyle.THICK);
            headerCellStyle.setBorderBottom(BorderStyle.THICK);
            //Border Color
            //We can use indexed colors
            headerCellStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            //Or any color we like using RGB Values
            //AWT color takes RED, GREEN, BLUE values

            XSSFColor customColor = new XSSFColor(new java.awt.Color(200, 10, 20));
            headerCellStyle.setLeftBorderColor(customColor);
            headerCellStyle.setRightBorderColor(customColor);
            headerCellStyle.setBottomBorderColor(customColor);
            /**
             * Fill Color and Pattern
             */
            headerCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());//needs to be set before fill background color
            headerCellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.LEAST_DOTS);

            /**
             * Fonts
             * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Font.html
             */
            Font headerFont = wb.createFont();
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setBold(true);
            headerFont.setFontName("Arial");
            headerFont.setColor(IndexedColors.BLACK.getIndex());
            //Assign the font to the style
            headerCellStyle.setFont(headerFont);
            /**
             * Text Angle (see also lecture 8)
             */
            headerCellStyle.setRotation((short) 45);

            /**
             * Shrink to fit
             * The text will be displayed smaller if it can't fit in the cell
             * In this case the font size that was assigned should be considered the maximum font size
             */
            headerCellStyle.setShrinkToFit(true);

            //Assign the style to all cells in first row
            for (Cell hc : sheet.getRow(0)) {
                hc.setCellStyle(headerCellStyle);
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
