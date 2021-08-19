package excel.export;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.*;

public class ExportExcelFromDatabase {
	
	public static void main(String args[]) throws Exception {
		
		Class.forName("com.mysql.cj.jdbc.Driver");
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/person", "root", "root");
        Statement statement=con.createStatement();
        ResultSet resultSet=statement.executeQuery("Select * from person");
        
        //ResultSet resultSet = statement.executeQuery("SELECT * from foo");
        ResultSetMetaData rsmd = resultSet.getMetaData();
        
        System.out.println(rsmd);
        
        int columnsNumber = rsmd.getColumnCount();
        System.out.println(columnsNumber);
        
        for (int i = 1; i <= columnsNumber; i++)
        System.out.print(" "+rsmd.getColumnName(i));
        
        System.out.println("\n");
        
        while (resultSet.next()) {
        	
            for (int i = 1; i <= columnsNumber; i++) {
            	
                if (i > 1) System.out.print(",  ");
                String columnValue = resultSet.getString(i);
                int getrow = resultSet.getRow();
                System.out.print(columnValue + " " + getrow/*+ rsmd.getColumnName(i)*/);
            }
            System.out.println("");
        }
	}

}

