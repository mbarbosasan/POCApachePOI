import org.apache.poi.hssf.record.formula.functions.Rows;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class Main {
    public static void main(String[] args) {

        ExcelUtil excelModel = new ExcelUtil("C:/temp/newExcel1.xls");
        try {
            excelModel.createExcelModel("UN-RIO");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

    }
}
