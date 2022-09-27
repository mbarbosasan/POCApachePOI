import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelUtil {
    private String path;
    private List<String> sampleGrandezas = new ArrayList<>();
    private List<String> sampleElementoEP = new ArrayList<>();
    private List<String> sampleUnidadeElementos = new ArrayList<>();

    public ExcelUtil(String path) {
        this.path = path;
    }

    public XSSFWorkbook createExcelModel(String unidadeNegocio) throws IOException, InvalidFormatException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Model");

        Row row1 = sheet.createRow(1);
        Cell cellA = row1.createCell(0);
        cellA.setCellValue("Cadastro de Variaveis em Batelada");

        Row row2 = sheet.createRow(2);
        cellA = row2.createCell(0);
        Cell cellB = row2.createCell(1);
        cellA.setCellValue("Data de Exportação");
        cellB.setCellValue(new Date());
        changeFormat(workbook, cellB);

        Row row3 = sheet.createRow(3);
        cellA = row3.createCell(0);
        cellB = row3.createCell(1);
        cellA.setCellValue("Unidade de Negocio");
        cellB.setCellValue(unidadeNegocio);

        Row row5 = sheet.createRow(5);
        cellA = row5.createCell(0);
        cellB = row5.createCell(1);
        Cell cellC = row5.createCell(2);
        Cell cellD = row5.createCell(3);
        Cell cellE = row5.createCell(4);
        Cell cellF = row5.createCell(5);
        Cell cellG = row5.createCell(6);
        Cell cellH = row5.createCell(7);
        Cell cellI = row5.createCell(8);

        cellA.setCellValue("Ação");
        setTableHeaderStyle(sheet, cellA);
        cellB.setCellValue("Elemento");
        setTableHeaderStyle(sheet, cellB);
        cellC.setCellValue("Grandeza");
        setTableHeaderStyle(sheet, cellC);
        cellD.setCellValue("Tipo");
        setTableHeaderStyle(sheet, cellD);
        cellE.setCellValue("Fonte");
        setTableHeaderStyle(sheet, cellE);
        cellF.setCellValue("Servidor");
        setTableHeaderStyle(sheet, cellF);
        cellG.setCellValue("Tag/Path/PI Fórmula");
        setTableHeaderStyle(sheet, cellG);
        cellH.setCellValue("Unidade na Fonte");
        setTableHeaderStyle(sheet, cellH);
        cellI.setCellValue("Resultado!");
        setTableHeaderStyle(sheet, cellI);

        Row row6 = sheet.createRow(6);
        cellA = row6.createCell(0);
        cellB = row6.createCell(1);
        cellC = row6.createCell(2);
        cellD = row6.createCell(3);
        cellE = row6.createCell(4);
        cellF = row6.createCell(5);
        cellG = row6.createCell(6);
        cellH = row6.createCell(7);
        cellI = row6.createCell(8);

        List<String> exampleRow = new ArrayList<>();
        exampleRow.add("INCLUIR");
        exampleRow.add("ELEMENTO_EP");
        exampleRow.add("GRANDEZA");
        exampleRow.add("Continua");
        exampleRow.add("Fonte");
        exampleRow.add("Servidor");
        exampleRow.add("TAG/PATH/PI Formula");
        exampleRow.add("Unidade na Fonte");
        exampleRow.add("PROCESSADA!");

        sampleGrandezas.addAll(exampleRow);
        sampleElementoEP.addAll(exampleRow);

        //Percorrendo todas as linhas independente se estao preenchidas ou nao para a formatacao da tabela.

        int rowStart = 6;
        int rowEnd = 50;

        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int i = 0; i <= 8; i++) {
                Cell currentCell = row.createCell(i);
                setTableCellStyle(sheet, currentCell);
                if (row.getRowNum() < 10) {
                    currentCell.setCellValue(exampleRow.get(i));
                    switch (currentCell.getColumnIndex()) {
                        case 1:
                            createDropDown(sheet, currentCell, sampleGrandezas);
                            break;
                        case 2:
                            createDropDown(sheet, currentCell, sampleElementoEP);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        getAllVariaveis(sheet);

        File file = new File(this.path);
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        return workbook;
    }

    public void setTableHeaderStyle(Sheet sheet, Cell cell) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        //Background
        cellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //Border
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);

        cell.setCellStyle(cellStyle);
        //Font color
        Font font = sheet.getWorkbook().createFont();
        font.setColor(HSSFColor.WHITE.index);
        cellStyle.setFont(font);
        sheet.autoSizeColumn(cell.getColumnIndex());
    }

    public void setTableCellStyle(Sheet sheet, Cell cell) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cell.setCellStyle(cellStyle);
        sheet.autoSizeColumn(cell.getColumnIndex());
    }

    public void changeFormat(XSSFWorkbook workbook, Cell cell) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("M/D/YYYY H:MM"));
        cell.setCellStyle(cellStyle);
    }

    public void createDropDown(Sheet sheet, Cell cell, List<String> list) {
        String arr [] = list.toArray(new String[list.size()]);
        DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dataValidationConstraint = dataValidationHelper.createExplicitListConstraint(arr);
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex());
        DataValidation validation = dataValidationHelper.createValidation(dataValidationConstraint, cellRangeAddressList);
        validation.setSuppressDropDownArrow(true);
        sheet.addValidationData(validation);
    }

    public void getAllVariaveis(Sheet sheet) {
        List<Variavel> listVariaveis = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() > 6) {
                    Variavel variavel = new Variavel();
                }
                listVariaveis.addAll(list);
            }
        }
        for (String item : listVariaveis) {
            System.out.println(item);
        }
    }
}
