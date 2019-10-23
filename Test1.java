package com.exceltoxml.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Test1 {

	public static void main(String[] args) throws ParserConfigurationException {
		FileInputStream inputStream;
		HashMap<String, List<Object[]>> map = new HashMap<String, List<Object[]>>();

		try {

			inputStream = new FileInputStream(new File("E:\\asn3.xls"));
			Workbook workbook = new HSSFWorkbook(inputStream);
			for (int i = 1; i < workbook.getNumberOfSheets(); i++) {

				map.put(workbook.getSheetName(i), getData(i, workbook));
			}
			HashMap<String,Object> baseMap=new HashMap<String, Object>();
			Sheet sheet = workbook.getSheet("baseXML");
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			while (rowIterator.hasNext()) {
              Row row=rowIterator.next();
              String node=row.getCell(0).getStringCellValue();
              String value=row.getCell(1).getStringCellValue();
              if(value.equals("root")) {
            	  baseMap.put(node,new HashMap<String,Object>());
            	  
              }
              else {
            	  String nodes[]=node.split("\\.");
            	  System.out.println(nodes.length);
            	  if(nodes.length==1) {
            		  HashMap<String,Object> innerMap=new HashMap<String, Object>();
            		innerMap=(HashMap<String, Object>) baseMap.get("companies");
            		  innerMap.put(nodes[0],new HashMap<String,Object>());
            		  System.out.println(innerMap);
            		  map.put("companies",(List<Object[]>) innerMap);
            		 
            	  }
            	  
              }
			}
         System.out.println(baseMap.values());
		} catch (

		FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static List<Object[]> getData(int i, Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(i);
		Iterator<Row> rowIterator = sheet.iterator();
		List<Object[]> list = new ArrayList<Object[]>();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Object[] object = new Object[row.getLastCellNum()];
			for (int j = 0; j < row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				object[j] = cell.getStringCellValue();

			}
			list.add(object);

		}
		return list;
	}

}
