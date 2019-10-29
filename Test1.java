package com.exceltoxml.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Test1 {
	static ArrayList<String> visited = new ArrayList<String>();
	static List<String> sheetList = new ArrayList<String>();
	static HashMap<String, Object> baseMap = new HashMap<String, Object>();
	static HashMap<String, List<Object[]>> map = new HashMap<String, List<Object[]>>();

	public static void main(String[] args) throws ParserConfigurationException, TransformerException {
		FileInputStream inputStream;
		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
		Document doc = docBuilder.newDocument();
		sheetList.add("company");
		sheetList.add("employee");
		try {

			inputStream = new FileInputStream(new File("E:\\asn3.xls"));
			Workbook workbook = new HSSFWorkbook(inputStream);
			for (int i = 1; i < workbook.getNumberOfSheets(); i++) {

				map.put(workbook.getSheetName(i), getData(i, workbook));
			}

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
						if (i > 0) {
							while (innerMap.containsKey(nodes[i - 1])) {
								// System.out.println("hi");
								innerMap = (HashMap<String, Object>) innerMap.get(nodes[i - 1]);

							}
						}

						if (sheetList.contains(nodes[i]) && !checkVisit(nodes[i])) {
							innerMap.put(nodes[i], new HashMap<String, Object>());
							visited.add(nodes[i]);
						} else {
							if (!sheetList.contains(nodes[i]))
								innerMap.put(nodes[i], "" + values[1]);
						}
						// System.out.println(baseMap.values());
					}

				}
			}
			Element rootElement = doc.createElement("Awes");
			doc.appendChild(rootElement);
			// System.out.println(baseMap.values());
			buildTags(doc, baseMap, rootElement);

			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			DOMSource source = new DOMSource(doc);
			StreamResult console = new StreamResult(System.out);
			transformer.transform(source, console);

		} catch (

		FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static void buildTags(Document doc, HashMap<String, Object> baseMap1, Element rootElement) {

		for (Entry<String, Object> entry : baseMap1.entrySet()) {
			String key = entry.getKey();

			Object value = entry.getValue();
			String str = "" + value.getClass();
			String ty[] = str.split("\\.");
			System.out.println(key + ":" + value);

			if (ty[ty.length - 1].contentEquals("HashMap")) {
				Element baseElement = doc.createElement(key);
				rootElement.appendChild(baseElement);
				buildTags(doc, (HashMap<String, Object>) value, baseElement);
			} else {
				Element element = doc.createElement(key);
				element.setTextContent((String) value);
				rootElement.appendChild(element);

			}

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
