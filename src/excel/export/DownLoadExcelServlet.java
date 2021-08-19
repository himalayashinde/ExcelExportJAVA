package excel.export;
import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
 
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 
// Extend HttpServlet class
public class DownLoadExcelServlet extends HttpServlet {
    private static final long serialVersionUID = 2067115822080269398L;
 
    public void init() throws ServletException {
        // Do nothing
    }
 
    public void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
 
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment; filename=SampleExcel.xls");
            
        //	response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
       //  response.setHeader("Content-Disposition","attachment; filename=report.xlsx"); 
        	
        	
            HSSFWorkbook workbook = createExcel();
            workbook.write(response.getOutputStream());
        } catch (Exception e) {
            throw new ServletException("Exception in DownLoad Excel Servlet", e);
        }
    }
 
    private HSSFWorkbook createExcel() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet worksheet = workbook.createSheet("My First POI Worksheet");
 
        HSSFRow row1 = worksheet.createRow(0);
 
        HSSFCell cellA1 = row1.createCell(0);
        cellA1.setCellValue("Hurray! You did it.");
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellA1.setCellStyle(cellStyle);
        return workbook;
    }
 
    public void destroy() {
        // do nothing.
    }
}