package com.demo.mavenproject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportGeneration {
	static String[] matchHeading = { "Xpath", "Source Content","Migrated Content", "Actual Style" };
	static String[] mismatchHeading = {"Xpath", "Source Content" };

	static String projectPath = System.getProperty("user.dir");

	XSSFWorkbook matchReport = new XSSFWorkbook();
	XSSFSheet matchRecords = matchReport.createSheet("matchRecords");
	XSSFWorkbook mismatchReport = new XSSFWorkbook();
	XSSFSheet mismatchRecords = mismatchReport.createSheet("mismatchRecords");

	public void generateMatchReport(ArrayList<String[]> list) throws IOException {
		FileOutputStream contentReport = new FileOutputStream(
				new File(projectPath + "//TestReports//MatchReport.xlsx"));
		Font headerFont = setHeaderFont(matchReport);

		XSSFCellStyle headerStyle = setHeaderStyle(matchReport, headerFont);

		Font ContentFont = setContentFont(matchReport);

		XSSFCellStyle ContentStyle = setContentStyle(matchReport, ContentFont);

		loadMatchHeaderToExcel(headerStyle);
		loadMatchContentToExcel(list, ContentStyle);
		matchRecords.setDefaultColumnWidth(50);
		matchRecords.autoSizeColumn(3);
		
		matchReport.write(contentReport);
		contentReport.close();
	}

	public void generatemisMatchReport(ArrayList<String[]> list) throws IOException {
		FileOutputStream contentReport = new FileOutputStream(
				new File(projectPath + "//TestReports//MisMatchReport.xlsx"));
		Font mismatchheaderFont = setHeaderFont(mismatchReport);
		XSSFCellStyle mismatchheaderStyle = setHeaderStyle(mismatchReport, mismatchheaderFont);
		Font ContentFont = setContentFont(mismatchReport);
		XSSFCellStyle ContentStyle = setContentStyle(mismatchReport, ContentFont);
		loadMisMatchHeaderToExcel(mismatchheaderStyle);
		loadmisMatchContentToExcel(list, ContentStyle);
		mismatchRecords.setDefaultColumnWidth(50);
		
		mismatchReport.write(contentReport);
		
	}

	private XSSFCellStyle setHeaderStyle(XSSFWorkbook report2, Font headerFont) {
		// TODO Auto-generated method stub
		XSSFCellStyle headerStyle = report2.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THICK);
		headerStyle.setBorderLeft(BorderStyle.THICK);
		headerStyle.setBorderRight(BorderStyle.THICK);
		headerStyle.setBorderTop(BorderStyle.THICK);
		headerStyle.setWrapText(true);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFont(headerFont);
		return headerStyle;
	}

	private Font setHeaderFont(XSSFWorkbook report2) {
		// TODO Auto-generated method stub
		Font headerFont = report2.createFont();
		headerFont.setBold(true);
		headerFont.setFontName("Arial");
		return headerFont;
	}

	private XSSFCellStyle setContentStyle(XSSFWorkbook report2, Font headerFont) {
		// TODO Auto-generated method stub
		XSSFCellStyle contentStyle = report2.createCellStyle();
		contentStyle.setBorderBottom(BorderStyle.THIN);
		contentStyle.setBorderLeft(BorderStyle.THIN);
		contentStyle.setBorderRight(BorderStyle.THIN);
		contentStyle.setBorderTop(BorderStyle.THIN);
		contentStyle.setWrapText(true);
		contentStyle.setAlignment(HorizontalAlignment.CENTER);
		contentStyle.setFont(headerFont);
		return contentStyle;
	}

	private Font setContentFont(XSSFWorkbook report2) {
		// TODO Auto-generated method stub
		Font contentFont = report2.createFont();
		contentFont.setFontName("Arial");
		return contentFont;
	}

	private void loadMatchHeaderToExcel(XSSFCellStyle headerStyle) {
		// TODO Auto-generated method stub
		Row row = matchRecords.createRow(0);
		XSSFCell cell = (XSSFCell) row.createCell(0);
		cell.setCellValue(matchHeading[0]);
		cell.setCellStyle(headerStyle);
		XSSFCell cell1 = (XSSFCell) row.createCell(1);
		cell1.setCellValue(matchHeading[1]);
		cell1.setCellStyle(headerStyle);
		XSSFCell cell2 = (XSSFCell) row.createCell(2);
		cell2.setCellValue(matchHeading[2]);
		cell2.setCellStyle(headerStyle);
		XSSFCell cell3 = (XSSFCell) row.createCell(3);
		cell3.setCellValue(matchHeading[3]);
		cell3.setCellStyle(headerStyle);

	}

	private void loadMisMatchHeaderToExcel(XSSFCellStyle headerStyle) {
		// TODO Auto-generated method stub
		Row row = mismatchRecords.createRow(0);
		XSSFCell cell = (XSSFCell) row.createCell(0);
		cell.setCellValue(mismatchHeading[0]);
		cell.setCellStyle(headerStyle);
		XSSFCell cell1 = (XSSFCell) row.createCell(1);
		cell1.setCellValue(mismatchHeading[1]);
		cell1.setCellStyle(headerStyle);
	}

	private void loadMatchContentToExcel(ArrayList<String[]> list, XSSFCellStyle ContentStyle) {
		// TODO Auto-generated method stub
		int rowNum = 1;
		for (String[] headerPara : list) {
			Row row = matchRecords.createRow(rowNum);
			XSSFCell cell = (XSSFCell) row.createCell(0);
			cell.setCellValue(headerPara[0]);
			cell.setCellStyle(ContentStyle);
			XSSFCell cell1 = (XSSFCell) row.createCell(1);
			cell1.setCellValue(headerPara[1]);
			cell1.setCellStyle(ContentStyle);
			XSSFCell cell2 = (XSSFCell) row.createCell(2);
			cell2.setCellValue(headerPara[2]);
			cell2.setCellStyle(ContentStyle);
			XSSFCell cell3 = (XSSFCell) row.createCell(3);
			cell3.setCellValue(headerPara[3]);
			cell3.setCellStyle(ContentStyle);
			rowNum++;
		}

	}

	private void loadmisMatchContentToExcel(ArrayList<String[]> list, XSSFCellStyle ContentStyle) {
		// TODO Auto-generated method stub
		int rowNum = 1;
		for (String[] headerPara : list) {
			Row row = mismatchRecords.createRow(rowNum);
			XSSFCell cell = (XSSFCell) row.createCell(0);
			cell.setCellValue(headerPara[0]);
			cell.setCellStyle(ContentStyle);
			XSSFCell cell1 = (XSSFCell) row.createCell(1);
			cell1.setCellValue(headerPara[1]);
			cell1.setCellStyle(ContentStyle);
			rowNum++;
		}

	}

}
