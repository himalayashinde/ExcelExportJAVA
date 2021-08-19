<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFSheet"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFWorkbook"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFCell"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFRow"%>
<%@ page import="org.apache.poi.hssf.usermodel.*"%>
<%@ page import="java.io.*" %>    
     
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Insert title here</title>
</head>
<body>

<%String name=request.getParameter("filename");%>
<%try{

HSSFWorkbook wb = new HSSFWorkbook();
HSSFSheet sheet = wb.createSheet("new sheet");
HSSFRow row = sheet.createRow((short)0);
HSSFCell cell = row.createCell((short)0);
cell.setCellValue(1);
row.createCell((short)1).setCellValue(1.2);
row.createCell((short)2).setCellValue("amkit Kumar ");
row.createCell((short)3).setCellValue(true);

FileOutputStream fileOut = new FileOutputStream("c:\\"+name+".xls");
wb.write(fileOut);
fileOut.close(); 
}catch ( Exception ex ){ 
} 
%>
Successfully created file.


</body>
</html>




