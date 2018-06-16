package be.quodlibet.be.apachepoi;

import java.util.Date;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/**
 *
 * @author Dries Horions <dries@quodlibet.be>
 */
public class ExcelUtils
{

    private static int xmlID = 0;//use a static variable to create unique id's
    /**
     * Get the value of a specific cell in the correct type
     * The returned value can be a String, Double or Date, depending on the cell type
     *
     * @param cell
     * @return
     */
    public static Object getTypedValue(Cell cell)
    {
        //Get the cell type
        CellType cellType = cell.getCellTypeEnum();

        if (cellType.equals(CellType.FORMULA)) {
            /**
             * For a formula cell, excel will store the last evaluated result
             * This can again be a value of any type (except formula)
             * For formula cells, we'll use the type of the result instead of the type of the cell
             */
            cellType = cell.getCachedFormulaResultTypeEnum();
            /**
             * The last calculated value is present in the cell, so we could simply use that
             * We can also recalculate the value using a formula evaluator
             */
            FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            //evaluateFormulaCell will update the cachedResult stored in the cell, and leave the formula untouched
            evaluator.evaluateFormulaCell(cell);
            //evaluator.evaluateInCell(cell) will remove the formula and replace it with the result, the formula can not be re-evaluated anymore
            //evaluator.evaluate(cell) will return the CellValue of the evaluation, but will not update the cell value

            /**
             * If we would want to know what the formula was that produced the result
             * we can get that like this :
             */
            String formula = cell.getCellFormula();
        }
        //Get the value of the cell based on the type of cell
        switch (cellType) {
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                /**
                 * Numeric cells can contain numbers, or dates
                 * the DateUtil can be used to check which one it is
                 */
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                else {
                    return cell.getNumericCellValue();
                }
            case STRING:
                return cell.getStringCellValue();
        }
        //If the cell was of no type we understand, return null (this shouldn't happen)
        return null;
    }

    /**
     * Columns are sized in 1/256th of a character
     * We'll create a simple method to size columns in a nr of characters
     * The maximum column size is 255 characters
     *
     * @param sheet
     * @param colIndex
     * @param charWidth
     */
    public static void setColSize(Sheet sheet, int colIndex, int charWidth)
    {
        if (charWidth <= 255) {
            sheet.setColumnWidth(colIndex, charWidth * 256);
        }
        else {
            sheet.setColumnWidth(colIndex, 255 * 256);
        }
    }

    /**
     * Angle the text in a cell at specific angle
     *
     * @param sheet
     * @param row
     * @param col
     * @param angle
     */
    public static void angleCell(Sheet sheet, int row, int col, int angle)
    {

        CellStyle angleStyle = sheet.getWorkbook().createCellStyle();
        angleStyle.cloneStyleFrom(sheet.getRow(row).getCell(col).getCellStyle());
        angleStyle.setRotation((short) angle);
        sheet.getRow(row).getCell(col).setCellStyle(angleStyle);
    }

    /**
     * @param sheet - The sheet in which to create the table
     * @param startRow - The row where the table should start (the row that will contain the column headers)
     * @param startColumn - The columns where the table should start
     * @param tStyle - The table style.
     * - A full list of styles is in this document "Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf" , you can download that file from
     * - this url : https://www.ecma-international.org/publications/standards/Ecma-376.htm ( ECMA-376 3rd edition Part 1 ) page 4970
     * - A full list will be provided with the course
     * @param tName - A unique name for the table in the workbook
     * @param headers - A list of Strings that will serve as the header column for the table
     * @param values - A list of a list of strings that contains the rows of the of the table body
     */
    public static void formatAsStringTable(Sheet sheet, int startRow, int startColumn, String tStyle, String tName, List<String> headers, List<List<String>> values)
    {
        //Create a region in the sheet that will contain the table
        AreaReference region = new AreaReference(
                new CellReference(startRow, startColumn),
                new CellReference(startRow + values.size(), startColumn + headers.size() - 1));
        //Create a XSSF table
        XSSFTable table = ((XSSFSheet) sheet).createTable();
        //Get a reference to the CT Table so we can access the predefined table formats
        CTTable cttable = table.getCTTable();
        CTTableStyleInfo tableStyle = cttable.addNewTableStyleInfo();
        //Set the predefined table style name
        tableStyle.setName(tStyle);
        //Set if we want alternating colors for the columns
        tableStyle.setShowColumnStripes(false);
        //Set if we want alternating colors for the rows
        tableStyle.setShowRowStripes(true);
        //Set the reference of the Style to the region we defined.  The reference should be a string representation of the region
        cttable.setRef(region.formatAsString());
        //Set the auto filter reference if we want auto filtering added to the columns
        cttable.addNewAutoFilter().setRef(region.formatAsString());
        //Set unique id's for the table, a table name can't contain spaces
        tName = tName.replace(" ", "_").toUpperCase();
        cttable.setDisplayName(tName);
        cttable.setName(tName);
        //each table needs a unique numeric ID in the xml that is greater than 0.  We use our global static variable xmlID
        cttable.setId(++xmlID);

        //Create the columns list in the table
        CTTableColumns columns = cttable.addNewTableColumns();
        columns.setCount(headers.size()); //define number of columns
        int colIndex = startColumn;
        //Add the data to the sheet

        //Create the header row
        XSSFRow headerRow = ((XSSFSheet) sheet).createRow(startRow);
        for (String header : headers) {
            XSSFCell localXSSFCell = headerRow.createCell(colIndex);
            localXSSFCell.setCellValue(header);
            //Add the column to the columns list of the table
            CTTableColumn column = columns.addNewTableColumn();
            column.setName(header);
            column.setId((colIndex++) + 1);//columns are not zero indexed
        }
        //Create the rows
        int rowIndex = startRow + 1;
        for (List<String> record : values) {
            XSSFRow row = ((XSSFSheet) sheet).createRow(rowIndex++);
            colIndex = startColumn;
            for (String col : record) {
                    XSSFCell localXSSFCell = row.createCell(colIndex++);
                    //Set the value
                    localXSSFCell.setCellValue(col);

            }
        }
    }

    public static void formatAsTable(Sheet sheet, int startRow, int startColumn, String tStyle, String tName, List<String> headers, List<List<Object>> values, HashMap<Integer, CellStyle> columnSstyles)
    {
        AreaReference region = new AreaReference(new CellReference(startRow, startColumn), new CellReference(startRow + values.size(), startColumn + headers.size() - 1));
        XSSFTable table = ((XSSFSheet) sheet).createTable();
        CTTable cttable = table.getCTTable();
        CTTableStyleInfo tableStyle = cttable.addNewTableStyleInfo();
        tableStyle.setName(tStyle);
        tableStyle.setShowColumnStripes(false);
        tableStyle.setShowRowStripes(true);
        cttable.setRef(region.formatAsString());
        cttable.addNewAutoFilter().setRef(region.formatAsString());
        tName = tName.replace(" ", "_").toUpperCase();
        cttable.setDisplayName(tName);
        cttable.setName(tName);
        cttable.setId(++xmlID);

        //Create the columns list in the table
        CTTableColumns columns = cttable.addNewTableColumns();
        columns.setCount(headers.size()); //define number of columns
        int colIndex = startColumn;
        XSSFRow headerRow = ((XSSFSheet) sheet).createRow(startRow);
        for (String header : headers) {
            XSSFCell localXSSFCell = headerRow.createCell(colIndex);
            localXSSFCell.setCellValue(header);
            CTTableColumn column = columns.addNewTableColumn();
            column.setName(header);
            column.setId((colIndex++) + 1);//columns are not zero indexed
            sheet.autoSizeColumn(colIndex);
        }
        /**
         * Set the values of the cell according to the type of object that was passed
         */
        //Create the rows
        int rowIndex = startRow + 1;
        for (List<Object> record : values) {
            XSSFRow row = ((XSSFSheet) sheet).createRow(rowIndex++);
            colIndex = startColumn;
            //for (Object col : record) {
            for (int colIdx = 0; colIdx < record.size(); colIdx++) {
                Object col = record.get(colIdx);
                if (col != null) {
                    XSSFCell localXSSFCell = row.createCell(colIndex++);

                    if (columnSstyles.containsKey(colIdx)) {
                        //Assign the style associated with this column to this cell
                        localXSSFCell.setCellStyle(columnSstyles.get(colIdx));
                    }
                    Boolean typeIdentified = false;
                    if (col instanceof Double) {
                        localXSSFCell.setCellValue((Double) col);
                        typeIdentified = true;
                    }
                    if (col instanceof Long) {
                        localXSSFCell.setCellValue((Long) col);
                        typeIdentified = true;
                    }
                    if (col instanceof String) {
                        String value = (String) col;
                        if (value.startsWith("=")) {
                            //assign the formula to the cell, remove the = sign, we only used it to identify a formula
                            localXSSFCell.setCellFormula(value.replaceFirst("=", ""));
                        }
                        else {
                            localXSSFCell.setCellValue(value);
                        }
                        typeIdentified = true;
                    }
                    if (col instanceof Date) {
                        localXSSFCell.setCellValue((Date) col);
                        typeIdentified = true;
                        // localXSSFCell.setCellStyle(dateStyle);
                    }
                    if (col instanceof Integer) {
                        localXSSFCell.setCellValue((Integer) col);
                        typeIdentified = true;
                    }
                    if (!typeIdentified) {
                        //if we couldn't identify the type of the object, put it's string representation in the cell
                        localXSSFCell.setCellValue(col.toString());
                    }
                }
            }
        }

    }

    public static void colorScaleRangeNumber(Sheet sheet, int startRow, int startCol, int endRow, int endCol, String lowColorRGB, String medianColorRGB, String highColorRGB)
    {
        //Get reference to Conditional Formatting rules of the sheet
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        //Create a range the formatting will be applied to
        CellRangeAddress[] region = {CellRangeAddress.valueOf(
            CellReference.convertNumToColString(startCol) + (startRow + 1) + ":"
            + CellReference.convertNumToColString(endCol) + (endRow + 1))};
        //Create a rule, in thios case a ConditionalFormattingColorScaleRule
        ConditionalFormattingRule rule
                                  = sheetCF.createConditionalFormattingColorScaleRule();

        //Set the ranges, assign a color to low, high and median value
        ColorScaleFormatting cs1 = rule.getColorScaleFormatting();
        cs1.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
        ((ExtendedColor) cs1.getColors()[0]).setARGBHex(lowColorRGB);
        cs1.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
        ((ExtendedColor) cs1.getColors()[2]).setARGBHex(highColorRGB);
        //percentile 50
        cs1.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        cs1.getThresholds()[1].setValue(50d);
        ((ExtendedColor) cs1.getColors()[1]).setARGBHex(medianColorRGB);
        //add the rule to the sheet
        sheetCF.addConditionalFormatting(region, rule);
    }
    public static void iconSetRange(Sheet sheet, int startRow, int startCol, int endRow, int endCol, IconSet iconSet)
    {
        //Get reference to Conditional Formatting rules of the sheet
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        //Create a range the formatting will be applied to
        CellRangeAddress[] region = {CellRangeAddress.valueOf(
            CellReference.convertNumToColString(startCol) + (startRow + 1) + ":"
            + CellReference.convertNumToColString(endCol) + (endRow + 1))};
        //Create a rule, in thios case a ConditionalFormattingColorScaleRule
        ConditionalFormattingRule rule
                                  = sheetCF.createConditionalFormattingRule(iconSet);
        IconMultiStateFormatting iconState  = rule.getMultiStateFormatting();
        iconState.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        iconState.getThresholds()[0].setValue(0.0);
        iconState.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        iconState.getThresholds()[1].setValue(25.0);
        iconState.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        iconState.getThresholds()[2].setValue(50.0);
        iconState.getThresholds()[3].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        iconState.getThresholds()[3].setValue(75.0);
        iconState.getThresholds()[4].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        iconState.getThresholds()[4].setValue(100.0);
        sheetCF.addConditionalFormatting(region, rule);
    }

    public static void dataBarRange(Sheet sheet, int startRow, int startCol, int endRow, int endCol, String colorRGB)
    {
        //Get reference to Conditional Formatting rules of the sheet
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        //Create a range the formatting will be applied to
        CellRangeAddress[] region = {CellRangeAddress.valueOf(
            CellReference.convertNumToColString(startCol) + (startRow + 1) + ":"
            + CellReference.convertNumToColString(endCol) + (endRow + 1))};
        //Create a color
        ExtendedColor color = sheet.getWorkbook().getCreationHelper().createExtendedColor();
        color.setARGBHex(colorRGB);
        //Create a rule passing the color
        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(color);
        //Assign the rule to the conditional formatting
        sheetCF.addConditionalFormatting(region, rule);
    }

}
