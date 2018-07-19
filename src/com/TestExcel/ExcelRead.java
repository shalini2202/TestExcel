package com.TestExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	
	
	public String getTestData() throws IOException {
		String name1 = null;
	ArrayList<String> ar = new ArrayList<String>();
	int flag = 0;
	
	String filePath = "C:\\Users\\a631020\\eclipse-workspace\\ExcelRead.xlsx";
	FileInputStream fis = new FileInputStream(filePath);
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	XSSFSheet sheet = wb.getSheet("Sheet1");
	XSSFRow row = sheet.getRow(0);
	
       int colNum = row.getLastCellNum();
       //System.out.println("Total Number of Columns in the  TestData.xlsx is : "+colNum);
       int rowNum = sheet.getLastRowNum()+1;
       //System.out.println("Total Number of Rows in the TestData.xlsx is : "+rowNum);
       
       for(int j=0; j<rowNum; j++) {
	    	
	    	for(int i=0; i<colNum; i++) {
				String name = sheet.getRow(j).getCell(i).getStringCellValue();
				ar.add(name);
				//System.out.println(name);
			}
  
	
       }
       
       Iterator itr=ar.iterator();  
       while(itr.hasNext()){ 
    	    name1 = (String) itr.next();
        System.out.println(name1);  
       } 
       
       fis.close();
	
	return name1;
}


	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/*
		String FilePath = "C:\\Users\\a631020\\eclipse-workspace\\ExcelRead.xlsx";
		FileInputStream fs = new FileInputStream(FilePath);
		XSSFWorkbook a =new XSSFWorkbook(fs);
		XSSFSheet sh = a.getSheet("Sheet1");

		XSSFRow row = sh.getRow(0);
		int colNum = row.getLastCellNum();
	    int rowNum = sh.getLastRowNum()+1;*/
		
		/*//Reading all the column in the row
		for(int i=0; i<colNum; i++) {
			String name = sh.getRow(0).getCell(i).getStringCellValue();
			System.out.println(name);
		}*/
		
		/*//Reading whole excel
	    for(int j=0; j<rowNum; j++) {
	    	
	    	for(int i=0; i<colNum; i++) {
				String name = sh.getRow(j).getCell(i).getStringCellValue();
				System.out.println(name);
			}
	    }*/
		
	    ExcelRead er = new ExcelRead();
	    er.getTestData();

}
}
