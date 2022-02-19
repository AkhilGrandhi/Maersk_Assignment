package com.testexcel.test;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;


public class App {

    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\AkhilGrandhi\\Downloads\\input1.xls";
    private static final String OUT_EXCEL_FILE_LOCATION = "C:\\Users\\AkhilGrandhi\\Downloads\\output.xls";
    static TreeMap<String, String> tree_map  = new TreeMap<String, String>();
    public static void main(String[] args) {

        Workbook workbook = null;
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            Cell cell1 =null;
            Cell cell2 =null;
            Sheet sheet = workbook.getSheet(0);
          
            for(int i = 1; i<sheet.getRows();i++) {
            	cell1= sheet.getCell(0, i); //column,row         
            	cell2= sheet.getCell(1, i);             
                try {
					tree_map.put(cell1.getContents(), cell2.getContents());
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
            
            writeExcel();
            
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }

    }
    
     public static void writeExcel() {
    	 WritableWorkbook myFirstWbook = null;
	        try {
	        	 WritableCellFormat wc =	 new WritableCellFormat();
    			 wc.setBackground(Colour.YELLOW);
	            myFirstWbook = Workbook.createWorkbook(new File(OUT_EXCEL_FILE_LOCATION));

	            // create an Excel sheet
	            WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);

	            // add something into the Excel sheet
	            Label label = new Label(0, 0, "Fruits");
	            excelSheet.addCell(label);     

	            label = new Label(1, 0, "Price per (Kg)");
	            excelSheet.addCell(label);
	            int i = 1;
	            
	            for (Map.Entry<String, String> entry : tree_map.entrySet()) {
	            			
	            		
	            		 if(Integer. parseInt(entry.getValue()) >= 50) {
	            			 label = new Label(0, i, entry.getKey(),wc);
		            		 excelSheet.addCell(label);
	            		  label = new Label(1, i, entry.getValue(),wc);
	            		  excelSheet.addCell(label);
	            		 }else {
	            			 label = new Label(0, i, entry.getKey());
		            		 excelSheet.addCell(label);
	            			 label = new Label(1, i, entry.getValue());
		            		  excelSheet.addCell(label);
	            		 }
	            		 i++;
	            }

	            myFirstWbook.write();


	        } catch (IOException e) {
	            e.printStackTrace();
	        } catch (WriteException e) {
	            e.printStackTrace();
	        } finally {

	            if (myFirstWbook != null) {
	                try {
	                    myFirstWbook.close();
	                } catch (IOException e) {
	                    e.printStackTrace();
	                } catch (WriteException e) {
	                    e.printStackTrace();
	                }
	            }


	        }

		
    	 
     }

}