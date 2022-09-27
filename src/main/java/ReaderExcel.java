import com.sun.corba.se.spi.orbutil.threadpool.Work;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ReaderExcel {

    private String path;

    public ReaderExcel(String path) {
        this.path = path;
    }

    public Workbook readExcel() throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream(new File(this.path));
        Workbook workbook = WorkbookFactory.create(fis);
        return workbook;
    }

    public void createDropdown(Sheet sheet) {
        CellRangeAddressList addressList = new CellRangeAddressList(0, 2, 0, 0);
        DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(new String[] {"Teste 1", "Teste 2", "Teste 3"});
    }
}
