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
import javax.xml.transform.TransformerConfigurationException;
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

public class ExceltoXml {
	static HashMap<String, Object> baseMap = new HashMap<String, Object>();
	static ArrayList<String> visited = new ArrayList<String>();
	static List<String> sheetList = new ArrayList<String>();

	private static boolean checkVisit(String str) {
		if (visited.contains(str))
			return true;
		return false;
	}

	public static void main(String[] args) {
		FileInputStream inputStream;
		try {

			inputStream = new FileInputStream(new File("E:\\test.xls"));
			Workbook workbook = new HSSFWorkbook(inputStream);
			String root = "";
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			Document doc = docBuilder.newDocument();
			// Storing the name of all the sheets
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetList.add(workbook.getSheetName(i));
			}
			Sheet sheet = workbook.getSheet("baseXML");
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			// This loop is used to construct the baseXml structure
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				String nodes[] = row.getCell(0).getStringCellValue().split("\\.");
				String values[] = row.getCell(1).getStringCellValue().split("\\.");
				String pkey = null;
				String primarykey[] = null;
				
				if (row.getCell(2) != null) {
					pkey = row.getCell(2).getStringCellValue();
					primarykey = pkey.split("\\.");
				}
				/*
				 * if (row.getCell(3) != null) { rkey = row.getCell(3).getStringCellValue();
				 * rekey = rkey.split("\\."); }
				 */
				if (values[0].equals("root")) {
					baseMap.put(nodes[0], new HashMap<String, Object>());
					root = nodes[0];
				} else {

					HashMap<String, Object> innerMap = (HashMap<String, Object>) baseMap.get(root);
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
							if (!sheetList.contains(nodes[i])) {
								innerMap.put(nodes[i], "" + values[1]);
								if ((pkey != null && !pkey.isEmpty())) {

									innerMap.put("pkey", primarykey[1]);
								}
							}
						}

					}

				}
			}

			// System.out.println(baseMap);
			Element rootElement = doc.createElement(root);
			doc.appendChild(rootElement);
			buildTags(doc, (HashMap<String, Object>) baseMap.get(root), rootElement, workbook, null);
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			DOMSource source = new DOMSource(doc);
			StreamResult console = new StreamResult(System.out);
			transformer.transform(source, console);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (TransformerConfigurationException e) {
			e.printStackTrace();
		} catch (TransformerException e) {
			e.printStackTrace();
		}
	}

	/**
	 * It recursively builds tag simultaneously reading the sheet
	 * 
	 * @param doc
	 * @param baseMap1
	 * @param rootElement
	 */
	private static void buildTags(Document doc, HashMap<String, Object> baseMap1, Element rootElement,
			Workbook workbook, String rkey) {

		for (Entry<String, Object> entry : baseMap1.entrySet()) {
			String key = entry.getKey();
			Object value = entry.getValue();
			String str = "" + value.getClass();
			String typeOfObject[] = str.split("\\.");
			if (typeOfObject[typeOfObject.length - 1].contentEquals("HashMap")) {
				if (sheetList.contains(key)) {
					HashMap<String, Object> hm = (HashMap<String, Object>) value;
					Sheet sheet = workbook.getSheet(key);
					Iterator<Row> rowIterator = sheet.iterator();
					rowIterator.next();

					while (rowIterator.hasNext()) {

						Row row = rowIterator.next();
						if (rkey == null || row.getCell(0).getStringCellValue().contentEquals(rkey)) {
							Element root1element = doc.createElement(key);
							rootElement.appendChild(root1element);
							for (Entry<String, Object> entry1 : hm.entrySet()) {
								String key1 = entry1.getKey();
								Object value1 = entry1.getValue();
								String str1 = "" + value1.getClass();
								String typeOfObject1[] = str1.split("\\.");
								if (typeOfObject1[typeOfObject1.length - 1].contentEquals("HashMap")) {
									String pkey = (String) hm.get("pkey");
									buildTags(doc, (HashMap<String, Object>) value, root1element, workbook,
											row.getCell(Integer.parseInt(pkey)).getStringCellValue());
								} else {
									// System.out.println(key1 + " : " + value1);
									if (key1.contentEquals("pkey"))
										continue;
									Cell cell = row.getCell(Integer.parseInt((String) value1));
									Element child = doc.createElement((String) key1);
									child.setTextContent(cell.getStringCellValue());
									root1element.appendChild(child);

								}
							}
						}
					}
				}
			}

		}
	}

}
