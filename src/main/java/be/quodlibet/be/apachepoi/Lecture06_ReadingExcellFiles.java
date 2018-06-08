package be.quodlibet.be.apachepoi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class Lecture06_ReadingExcellFiles
{
    public static void main(String[] args)
    {
        //Location where we will store the Excell Files used in this course
        //You can use any existing folder, for this lesson, ensure the example.xlsx file is there.
        //This file will be added as a resource to the current lecture.
        String excellFolder = "D:\\Udemy\\Projects\\ApachePOICourse\\resources\\";
        //The DataFormatter class can be used to format Excell data as a String exactly as it appears in the excell file
        DataFormatter dataFormatter = new DataFormatter();
        //Read the Excell file into an input stream
        try (InputStream inp = new FileInputStream(excellFolder + "example.xlsx")) {
            //Create a Workbook from the input stream
            Workbook wb = WorkbookFactory.create(inp);

            // Loop through all available sheets
            for (Sheet s : wb) {
                System.out.println(s.getSheetName());
            }

            //Now we'll read the content of the first sheet and print it
            Sheet sheet = wb.getSheetAt(0);

            //Loop through all rows of this sheet
            /**
             * you can also loop through the rows with a foreach loop
             * We'll use a for loop later in this lecture, because we want to also loop over the empty cells,
             * The foreach loop will skip these empty cells
             */
            for (Row rw : sheet) {
                for (Cell cll : rw) {
                    //If we uncomment below lines, we'll see that empty cells like H203 are not shown in the output
                    //System.out.print(cll.getAddress().formatAsString() + "\t");
                }
                //System.out.print("\n");
            }

            for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
                //Get a reference to the row
                Row row = sheet.getRow(r);
                /**
                 * Loop through all the columns in the row from the first non-empty cell to the last non-empty cell
                 * getFirstCellNum will return the first cell, index 0
                 * getLastCellNum will return the last cell PLUS ONE !!!
                 * in our case we have cells 0 - 11 populated,
                 * getLastCellNum will return 12
                 */
                for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                    //Get a reference to the cell
                    Cell cell = row.getCell(c);
                    /**
                     * It is possible that an excel file contains empty cells
                     * the getCell() method returns a null value in that case.
                     * Without checking explicitly for null values, our code would fail when encountering an empty cell
                     * In our case for example cell H203
                     */
                    String output = "";
                    if (cell == null) {
                        output = "EMPTY";
                    }
                    else {
                        /**
                         * Get the cell value using the DataFormatter
                         * this formats Excel data as a String exactly as it appears in the excel file
                         * for formula's this will return the string representation of the formula
                         */

                        String cellValue = dataFormatter.formatCellValue(cell);
                        /**
                         * If you need the cell data in the correct dataType, you have to check what type of cell you're dealing with
                         * we'll create a method getTypedValue in a new class called ExcelUtils
                         */
                        Object typedCellValue = ExcelUtils.getTypedValue(cell);
                        output = "" + typedCellValue;
                    }
                    //For easier understanding of console output, we format each value as a 20 character string
                    System.out.print(String.format("%-20.20s", output) + " | ");
                }
                //Print a new line
                System.out.print("\n");
            }
        }
        catch (IOException ex) {
            System.out.println("The file could not be read : " + ex.getMessage());
        }
            catch (InvalidFormatException ex) {
            System.out.println("The file format is not valid : " + ex.getMessage());
        }
        catch (EncryptedDocumentException ex) {
             System.out.println("The excell file is encrypted : " + ex.getMessage());
        }
    }


}
