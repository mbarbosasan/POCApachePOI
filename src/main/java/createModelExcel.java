import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Instant;
import java.util.Date;

public class createModelExcel {
    public static void main(String[] args) {
        HSSFWorkbook modelExcel = new HSSFWorkbook();
        HSSFSheet inputExcel = modelExcel.createSheet("Model Excel");
        Row row2 = inputExcel.createRow(2);
        Row row3 = inputExcel.createRow(3);
        Row row5 = inputExcel.createRow(5);
        CellStyle cellStyle = inputExcel.getWorkbook().createCellStyle();

        Cell cellA = row2.createCell(0);
        Cell cellB = row2.createCell(1);
        Cell cellC = row5.createCell(2);
        Cell cellD = row5.createCell(3);
        Cell cellE = row5.createCell(4);
        Cell cellF = row5.createCell(5);
        Cell cellG = row5.createCell(6);
        Cell cellH = row5.createCell(7);
        Cell cellI = row5.createCell(8);

        cellA.setCellValue("Data do modelo: ");
        cellB.setCellValue(Date.from(Instant.now()).toString());
        cellA = row3.createCell(0);
        cellB = row3.createCell(1);
        cellA.setCellValue("Unidade de negocio");
        cellB.setCellValue("UN-BS");
        cellA = row5.createCell(0);
        cellA.setCellValue("Acao");
        cellB.setCellValue("Elemento");
        cellC.setCellValue("Grandeza");
        cellD.setCellValue("Tipo");
        cellE.setCellValue("Fonte");
        cellF.setCellValue("Servidor");
        cellG.setCellValue("Tag/Path/PI Formula");
        cellH.setCellValue("Unidade na Fonte");
        cellI.setCellValue("Data - Resultado do Processamento");
        inputExcel.autoSizeColumn(0);
        inputExcel.autoSizeColumn(1);
        File file = new File("C:/temp/model.xls");
        FileOutputStream fo = null;
        try {
            fo = new FileOutputStream(file);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            modelExcel.write(fo);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

}
