package practice;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class UnderstandingExcelConnection {



    @Test
    public void testCase1() throws IOException {
        String filePath = "UnderstandingExcelConnection.xlsx";

        FileInputStream byteCodeExcelFile = new FileInputStream(filePath);

        Workbook workbook = WorkbookFactory.create(byteCodeExcelFile);

        Sheet workSheet = workbook.getSheet("Sheet1");

        Cell cell;
        cell = workSheet.getRow(0).getCell(0);
        System.out.println(cell.toString());

        Cell cell2 = workSheet.getRow(0).getCell(1);
        System.out.println(cell2);

        Cell cell3 = workSheet.getRow(0).getCell(2);

        if (cell3 == null)
            System.out.println("Cell is Empty");
        else
            System.out.println(cell3);


    }



}
