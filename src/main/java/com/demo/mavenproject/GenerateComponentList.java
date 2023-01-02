package com.demo.mavenproject;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenerateComponentList {
	public static String projectPath = System.getProperty("user.dir");
	public static List<String[]> table = new ArrayList<String[]>();
	public static List<String[]> para = new ArrayList<String[]>();
	public static void readExcel(String fileName) {
		try {

			FileInputStream file = new FileInputStream(new File(projectPath + "//MappingFile//" + fileName));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			// Row headerRow=rowIterator.next();
			while (rowIterator.hasNext()) {

				Row row = rowIterator.next();
				String[] entry = new String[3];
				for (int i = 0; i < 3; i++) {
					Cell str = row.getCell(i);

					if (!(str == null)) {
						System.out.print(str.toString() + " ");
						entry[i] = str.toString();
					} else
						entry[i] = "nullContent";
				}
				System.out.println();
				if ((entry[2].equals("Table")))
					table.add(entry);
				else if ((entry[2].equals("Paragraph")))
					para.add(entry);
			}
			System.out.println(table);
			System.out.println("++++++++++++++++++++++++++++++++++++++");
			System.out.println(para);
			file.close();
			workbook.close();
			/*
			 * int n=table.size(); System.out.println("file closed"+n); for(int i=0;i<n;i++)
			 * { String[] entry=table.get(i); for(int j=0;j<3;j++) {
			 * System.out.print(entry[j]+" *  "); } System.out.println(); }
			 */

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
