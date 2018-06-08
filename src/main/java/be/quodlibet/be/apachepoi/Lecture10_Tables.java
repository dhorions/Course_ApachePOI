package be.quodlibet.be.apachepoi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture10_Tables
{
    public static void main(String[] args)
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //Create an output stream to write the file
        String filePath = excellFolder + "lecture10.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            //Create a workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            //Create a sheet (the name can't contain ":")
            Sheet sheet = wb.createSheet("First Sheet");
            /**
             * Create a List of column headers
             */
            List<String> headers = Arrays.asList("Column 1", "Column 2", "Column 3", "Column 4");
            /**
             * For each row we'll also create a list of Strings, and we'll add all these lists (rows) to a big list that represents the table body
             */
            List<List<String>> data = new ArrayList();
            //Add 10 rows to the table body
            for (int i = 0; i < 10; i++) {
                data.add(Arrays.asList("Row " + i + " Column 1", "Row " + i + " Column 2", "Row " + i + " Column 3", "Row " + i + " Column 4"));
            }
            /**
             * In our ExcelUtils class, we will create a method that can take the list of column headers
             * and the table body list and format that as a table in Excel
             * It takes the following parameters
             * sheet - The sheet in which to create the table
             * startRow - The row where the table should start (the row that will contain the column headers)
             * startColumn - The columns where the table should start
             * tStyle - The table style.
             * - A full list of styles is in this document "Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf" , you can download that file from
             * - this url : https://www.ecma-international.org/publications/standards/Ecma-376.htm ( ECMA-376 3rd edition Part 1 ) page 4970
             * - A full list will be provided with the course
             * tName - A unique name for the table in the workbook
             * headers - A list of Strings that will serve as the header column for the table
             * values - A list of a list of strings that contains the rows of the of the table body
             */
            ExcelUtils.formatAsStringTable(sheet, 0, 0, "TableStyleDark1", "TableOne", headers, data);
            //Let's create some more tables to see some different style examples in our output
            List<String> tableStyles = Arrays.asList("TableStyleDark2", "TableStyleDark3", "TableStyleDark11", "TableStyleLight2", "TableStyleLight3", "TableStyleLight11", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium11");
            int startRow = 0;
            int tableNameIndex = 0;
            for (String tStyleName : tableStyles) {
                startRow += data.size() + 1;//start the next table the row after the previous table
                String tableName = "Table_" + ++tableNameIndex;//Give the table a unique name
                ExcelUtils.formatAsStringTable(sheet, startRow, 0, tStyleName, tableName, headers, data);
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
