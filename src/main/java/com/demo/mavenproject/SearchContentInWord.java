package com.demo.mavenproject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class SearchContentInWord {

	static List<XWPFTable> tableList;
	static List<XWPFParagraph> paralist;
	static int tableListSize = 0;
	public static ArrayList<String[]> matchData = new ArrayList<String[]>();
	public static ArrayList<String[]> mismatchData = new ArrayList<String[]>();;
	public static String projectPath = System.getProperty("user.dir");

	public SearchContentInWord() throws Exception {
		File inputFile = new File(projectPath + "\\MigratedDocx\\Table.docx");
		FileInputStream fiS = new FileInputStream(inputFile);
		XWPFDocument migDocx = new XWPFDocument(fiS);
		tableList = migDocx.getTables();
		paralist = migDocx.getParagraphs();
		tableListSize = tableList.size();
	}

	public void readTableData() {

		for (String[] tableContent : GenerateComponentList.table) {
			boolean match = false;
			String paraText = "";
			String Style = "";
			for (XWPFTable tbl : tableList) {
				int rows = tbl.getNumberOfRows();
				for (int a = 0; a < rows; a++) {
					List<XWPFTableCell> cell = tbl.getRow(a).getTableCells();
					for (XWPFTableCell c : cell) {
						List<XWPFParagraph> tablepara = c.getParagraphs();
						for (XWPFParagraph para : tablepara) {

							if (para.getText().contains(tableContent[1])) {
								paraText = para.getText();
								Style = para.getStyle();
								System.out.println(paraText + " -> " + Style);
								match = true;
							} 
						}
					}
				}
			}
			if (match) {
				matchData.add(new String[] { tableContent[0], tableContent[1], paraText, Style });
			} else {
				mismatchData.add(new String[] { tableContent[0], tableContent[1] });
			}
		}
	}

	public void readParaData() {
		for (String[] paraContent : GenerateComponentList.para) {
			boolean match = false;
			String paraText = "";
			String Style = "";
			for (XWPFParagraph para : paralist) {
				if (para.getText().contains(paraContent[1])) {
					paraText = para.getText();
					Style = para.getStyle();
					System.out.println(paraText + " -> " + Style);
					match = true;
				} 
			}
			if (match) {
				matchData.add(new String[] { paraContent[0], paraContent[1], paraText, Style });
			} else {
				mismatchData.add(new String[] { paraContent[0], paraContent[1] });
			}
		}
	}
}