package excel.export;
import  java.io.*;  
import  java.sql.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;  
import  org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import  org.apache.poi.hssf.usermodel.HSSFRow;
import  org.apache.poi.hssf.usermodel.HSSFCell;  
import javax.servlet.*;  
import javax.servlet.http.*;  


public class CreateExcelFile extends HttpServlet{
	
	public CreateExcelFile() {
		
	}
	
public void doPost(HttpServletRequest request,HttpServletResponse response ) {
	
//public static void main(String[]args){
try{
	response.setContentType("application/vnd.ms-excel");
    response.setHeader("Content-Disposition", "attachment; filename=DataExported.xls");
    
String filename="g:/DataExported.xls" ;
HSSFWorkbook hwb=new HSSFWorkbook();
HSSFSheet sheet =  hwb.createSheet("new sheet");

HSSFRow rowhead=   sheet.createRow((short)0);
rowhead.createCell((short)0).setCellValue("SNo");
rowhead.createCell((short)1).setCellValue("Name");
rowhead.createCell((short)2).setCellValue("Address");
rowhead.createCell((short)3).setCellValue("Contact No");
rowhead.createCell((short)4).setCellValue("E-mail");

Class.forName("com.mysql.cj.jdbc.Driver");
Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/test", "root", "root");
Statement st=con.createStatement();
ResultSet rs=st.executeQuery("Select * from employee");
request.getContentType();
request.getRequestDispatcher("Success2.jsp").forward(request,response);
int i=1;
while(rs.next()){
HSSFRow row=   sheet.createRow((short)i);
row.createCell((short)0).setCellValue(Integer.toString(rs.getInt("id")));
row.createCell((short)1).setCellValue(rs.getString("name"));
row.createCell((short)2).setCellValue(rs.getString("address"));
row.createCell((short)3).setCellValue(Double.toString(rs.getDouble("contactNo")));
row.createCell((short)4).setCellValue(rs.getString("email"));
i++;
}
FileOutputStream fileOut =  new FileOutputStream(filename);
hwb.write(fileOut);
fileOut.close();

//hwb.write(response.getOutputStream());

response.setContentType(filename);
System.out.println("Your excel file has been generated!");

} catch ( Exception ex ) {
    System.out.println(ex);

}
    }
}

//}