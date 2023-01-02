package com.demo.mavenproject;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

public class GetXpathsAndContent extends DefaultHandler {

    private String xPath = "/";
    private XMLReader xmlReader;
    private GetXpathsAndContent parent;
    private StringBuilder characters = new StringBuilder();
    public static ArrayList<String[]> xpathContent = new ArrayList<String[]>();
    private Map<String, Integer> elementNameCount = new HashMap<String, Integer>();
    public static ArrayList<String[]> xmlDetails = new ArrayList<String[]>();
    static int rowCount =1;
    
   // public HashMap<String>
    public GetXpathsAndContent(XMLReader xmlReader) {
        this.xmlReader = xmlReader;
    }

    private GetXpathsAndContent(String xPath, XMLReader xmlReader, GetXpathsAndContent parent) {
        this(xmlReader);
        this.xPath = xPath;
        this.parent = parent;
    }
    @Override
    public void startElement(String uri, String localName, String qName, Attributes atts) throws SAXException {
        Integer count = elementNameCount.get(qName);
        if(null == count) {
            count = 1;
        } else {
            count++;
        }
        elementNameCount.put(qName, count);
        String childXPath = xPath + "/" + qName + "[" + count + "]";
        GetXpathsAndContent child = new GetXpathsAndContent(childXPath, xmlReader, this);
        xmlReader.setContentHandler(child);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        String value = characters.toString().trim();
        
        if(value.length() > 0) {
           // System.out.println(xPath + "='" + characters.toString() + "'");
            xpathContent.add(new String[]{xPath, characters.toString()} );
        }
        xmlReader.setContentHandler(parent);
    }
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        characters.append(ch, start, length);
    }
    
    public static Map<String, List<String>> loadingComponent(XSSFSheet sheet, Map<String, List<String>> dataMap) {
    	for(int i=1;i<sheet.getLastRowNum()+1;i++) {
    		if(dataMap.containsKey(sheet.getRow(i).getCell(0).getStringCellValue())) {
    			dataMap.get(sheet.getRow(i).getCell(0).getStringCellValue()).add(sheet.getRow(i).getCell(1).getStringCellValue());
    		}else {
    			dataMap.put(sheet.getRow(i).getCell(0).getStringCellValue(), new ArrayList<String>(Arrays.asList(sheet.getRow(i).getCell(1).getStringCellValue())));
    		}
    	}
    	return(dataMap);
    }
    
    public static void addingContentToExcel(XSSFSheet spreadsheet, Map<String, List<String>> componentMap) {
    	 for(String[] xpathData :GetXpathsAndContent.xpathContent) {
         	System.out.println("Xpath is "+xpathData[0]);
         	spreadsheet.createRow(rowCount).createCell(0).setCellValue(xpathData[0]);
         	spreadsheet.getRow(rowCount).createCell(1).setCellValue(xpathData[1]);
         	boolean compFound = false;
         	for(String key : componentMap.keySet()) {
         		if(xpathData[0].contains(key)) {
         			spreadsheet.getRow(rowCount).createCell(2).setCellValue(componentMap.get(key).get(0));
         			compFound =true;
         			rowCount++;
         			break;
         		}else {
         			continue;
         		}
         	}if(!compFound) {
         	rowCount++;
         	}
         }
    }
   
}