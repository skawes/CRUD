package com.exceltoxml.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Test1 {
	static ArrayList<String> visited = new ArrayList<String>();

	public static void main(String[] args) throws ParserConfigurationException {
		FileInputStream inputStream;
		HashMap<String, List<Object[]>> map = new HashMap<String, List<Object[]>>();
		List<String> sheetList = new ArrayList<String>();
		sheetList.add("company");
		sheetList.add("employee");
		try {

			inputStream = new FileInputStream(new File("E:\\asn3.xls"));
			Workbook workbook = new HSSFWorkbook(inputStream);
			for (int i = 1; i < workbook.getNumberOfSheets(); i++) {

				map.put(workbook.getSheetName(i), getData(i, workbook));
			}
			HashMap<String, Object> baseMap = new HashMap<String, Object>();
			Sheet sheet = workbook.getSheet("baseXML");
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				String nodes[] = row.getCell(0).getStringCellValue().split("\\.");
				String values[] = row.getCell(1).getStringCellValue().split("\\.");
				if (values[0].equals("root")) {
					baseMap.put(nodes[0], new HashMap<String, Object>());

				} else {

					HashMap<String, Object> innerMap = (HashMap<String, Object>) baseMap.get("companies");
					for (int i = 0; i < nodes.length; i++) {

						while (innerMap.containsKey(nodes[nodes.length - 2])) {
							System.out.println("hi");
							innerMap = (HashMap<String, Object>) innerMap.get(nodes[nodes.length - 2]);

						}

						if (sheetList.contains(nodes[i]) && !checkVisit(nodes[i])) {
							innerMap.put(nodes[i], new HashMap<String, Object>());
							visited.add(nodes[i]);
						} else {
							if (!sheetList.contains(nodes[i]))
								innerMap.put(nodes[i], "" + 1);
						}
						System.out.println(baseMap.values());
					}

				}
			}

			/*
			 * for(Object set:baseMap.values()) { HashMap<String, Object>
			 * hm=(HashMap<String, Object>) set; System.out.println(hm.keySet()); }
			 */
		} catch (

		FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static boolean checkVisit(String str) {
		if (visited.contains(str))
			return true;
		return false;
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
