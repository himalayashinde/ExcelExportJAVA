<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<%@ page import="org.apache.poi.hssf.usermodel.*"%>
<%@ page import="org.apache.poi.ss.usermodel.*"%>
<%@ page import="java.io.*" %>    
    
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Success Exporting the Excel file</title>

<script>  

function sendInfo()  
{  
	alert("Your File is ready to dowloaded !!!");  
}  
  
</script> 

</head>
<body>

<!-- <form name="DownloadExcel" action="CreateExcelFile" method="POST">
<button type="submit" onClick="sendInfo()">Download Excel</button>
</form> -->
<!-- 
<form name="DownloadExcel" action="DownLoadExcelServlet" method="GET">
<button type="submit" onClick="sendInfo()">Download Excel</button>
</form>  -->

 <form name="DownloadExcel" action="DownLoadExcelServlet2" method="GET">
<button type="submit">Download Excel</button>
</form>

</body>
</html>




