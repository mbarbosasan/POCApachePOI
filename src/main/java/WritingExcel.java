import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WritingExcel {
    public static void main(String[] args) {
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet0 = workbook.createSheet("First sheet");

        Row row0 = sheet0.createRow(0);
//
//        Cell cellA = row0.createCell(0);
//        Cell cellB = row0.createCell(1);
//
//        cellA.setCellValue("First cell!");
//        cellB.setCellValue("Second cell!");

        for (int rows = 0; rows < 10; rows++) {
            Row row = sheet0.createRow(rows);
            for(int cols = 0; cols < 10; cols++) {
                Cell cell = row.createCell(cols);
                cell.setCellValue((int) (Math.random() * 100));
            }
        }

        File f = new File("C:/temp/teste.xls");
        FileOutputStream fo = null;
        try {
            fo = new FileOutputStream(f);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook.write(fo);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
