package com.shiv;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Create_Write_in_Excel {

	public static void main(String[] args) {
		//create a Blank WorkBook
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		//create a Blank Sheet
		HSSFSheet sheet = workbook.createSheet("Team Data");
		
		//This data needs to be written (Object[])
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		
		data.put("1", new Object[]{"Name" , "Area" , "EMail" , "Age"});
		data.put("2", new Object[]{"Shiv Kumar Napit" , "Developer" , "shivkumar.napit@trafigura.com", 26});
		data.put("3", new Object[]{"Gaurav Mulye" , "Developer" , "gaurav.mulye@trafigura.com",30});
		data.put("4", new Object[]{"Anuj Sharma" , "TeamLead" , "anuj.sharma@trafigura.com",29});
		data.put("5", new Object[]{"Gaurav Agarwal" , "Manager" , "gaurav.agarwal@trafigura.com",45});
		data.put("6", new Object[]{"Amarjit Kumar" , "Support Analyst" , "amarjit.kumar@trafigura.com",26});
		data.put("7", new Object[]{"Lalit Singh Rajput" , "Support Analyst" , "lalitsingh.rajput@trafigura.com",26});
		
		//Iterate over the data and Write to the file sheet
		
		Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        
        try{
        	//Write the workbook in file system
        	
        	FileOutputStream out = new FileOutputStream(new File("AlfrescoTeam.xls"));
        	workbook.write(out);
        	out.close();
        	
        	System.out.println("AlfrescoTeam.xls written successfully on disk");
        	
        }catch(Exception e){
        	e.printStackTrace();
        } 
	}
}
