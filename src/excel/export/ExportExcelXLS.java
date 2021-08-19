package excel.export;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * This program read date values from XLSX file in Java using Apache POI.
 * 
 * @author WINDOWS 8
 *
 */
public class ExportExcelXLS{

    public static void main(String[] args) throws FileNotFoundException,
                                        IOException {
        writeIntoExcel("birthdays.xlsx");
        readFromExcel("birthdays.xlsx");
    }
    
    public static void readFromExcel(String file) throws IOException{
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Birthdays");
        XSSFRow row = myExcelSheet.getRow(0);
        
        if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
            String name = row.getCell(0).getStringCellValue();
            System.out.println("NAME : " + name);
        }
        
        if(row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
            Date birthdate = row.getCell(1).getDateCellValue();
            System.out.println("DOB :" + birthdate);
        }
        
        myExcelBook.close();
        
    }

    @SuppressWarnings("deprecation")
    public static void writeIntoExcel(String file) throws 
                      FileNotFoundException, IOException{
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Birthdays");
        Row row = sheet.createRow(0); 

        Cell name = row.createCell(0);
        name.setCellValue("Gokul");
        
        Cell birthdate = row.createCell(1);
        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);

        birthdate.setCellValue(new Date(115, 10, 10));
        
        sheet.autoSizeColumn(1);
        
        book.write(new FileOutputStream(file));
        book.close();
    }
}

//NAME : Gokul
//DOB :Tue Nov 10 00:00:00 GMT+08:00 2015