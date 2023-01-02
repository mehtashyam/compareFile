package com.demo.mavenproject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class Demo {

	static XSSFWorkbook workbook = new XSSFWorkbook();
	public static  XSSFSheet spreadsheet = workbook.createSheet("Xpath Data");
    public static Map<String,List<String>> componentMap=new HashMap<>();
    public static String projectPath = System.getProperty("user.dir");
    
    public static void main(String[] args) throws Exception {
    	FileInputStream file = new FileInputStream(new File(projectPath +"//Components//Properties.xlsx"));
    	XSSFWorkbook dataworkbook = new XSSFWorkbook(file);
    	XSSFSheet sheet = dataworkbook.getSheetAt(0);
    	componentMap = GetXpathsAndContent.loadingComponent(sheet, componentMap);
    	System.out.println(componentMap);
        SAXParserFactory spf = SAXParserFactory.newInstance();
        SAXParser sp = spf.newSAXParser();
        XMLReader xr = sp.getXMLReader();
        xr.setContentHandler(new GetXpathsAndContent(xr));
        xr.parse(new InputSource(new FileInputStream(projectPath +"//SourceXMLs//table.xml")));
        spreadsheet.createRow(0).createCell(0).setCellValue("Xpath");
        spreadsheet.getRow(0).createCell(1).setCellValue("Xpath Content");
        spreadsheet.getRow(0).createCell(2).setCellValue("Component");
        GetXpathsAndContent.addingContentToExcel(spreadsheet, componentMap);
        File f =new File(projectPath+"\\MappingFile\\XpathData.xlsx");
        String fileName = f.getName();
        FileOutputStream out = new FileOutputStream(f);
        workbook.write(out);
        dataworkbook.close();
        out.close();
        GenerateComponentList.readExcel(fileName);
        SearchContentInWord search = new SearchContentInWord();
        search.readParaData();
        search.readTableData();
        ReportGeneration rg = new ReportGeneration();
        rg.generateMatchReport(SearchContentInWord.matchData);
        rg.generatemisMatchReport(SearchContentInWord.mismatchData);
    }
}