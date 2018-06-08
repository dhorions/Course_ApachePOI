package be.quodlibet.be.apachepoi;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateCustomPivotTable
{

    public static void main(String[] args) throws IOException, InvalidFormatException
    {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();

        //Create some data to build the pivot table on
        setCellData(sheet);

        new CustomPivotTable(sheet, "A1:D6", "F3")
                .addRowLabel(0) // set first column as 1-th level of rows
                .excludeSubTotal(0) // excude subtotal
                .addRowLabel(1) // set second column of source as 2-th level of rows
                .addColLabel(3)
                .setFormatPivotField(3, 9)
                .addColumnLabel(DataConsolidateFunction.SUM, 2) // Sum up the second column
                .setFormatDataField(2, 4); //# ##0.00

        FileOutputStream fileOut = new FileOutputStream("custom-pivottable.xlsx");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }

    private static void setCellData(XSSFSheet sheet)
    {

        String[] cities = {"Rome", "Paris", "Rome", "Paris", "Athens"};
        String[] names = {"Jane", "Tarzan", "Terk", "Kate", "Dmitry"};
        Integer[] balances = {107634, 554234, 10234, 22350, 15234};
        Double[] percents = {0.25, 0.5, 0.75, 0.25, 0.5};

        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("City");
        row.createCell(1).setCellValue("Name");
        row.createCell(2).setCellValue("Balance");
        row.createCell(3).setCellValue("Percents");

        for (int i = 0; i < cities.length; i++) {
            row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(cities[i]);
            row.createCell(1).setCellValue(names[i]);
            row.createCell(2).setCellValue(balances[i]);
            row.createCell(3).setCellValue(percents[i]);
        }
    }
}

