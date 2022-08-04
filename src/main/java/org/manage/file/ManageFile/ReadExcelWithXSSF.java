package org.manage.file.ManageFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hpsf.Vector;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RangeCopier;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.cellwalk.CellWalk;
import org.apache.poi.xssf.usermodel.XSSFRangeCopier;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelWithXSSF {

	public static void main(String[] args) throws IOException {
		// Read XSL file
        FileInputStream inputStream = new FileInputStream(new File("C:Users/csra/Desktop/Cheque.xlsx"));

        // Get the workbook instance for XLS file
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        for (int i = 1; i < workbook.getNumberOfSheets(); i++)
        {
        	 //active sheet
        	 workbook.setActiveSheet(i);
        	 //get active sheet
        	 XSSFSheet sourceSheet =  workbook.getSheetAt(i);
        	 //get lastRowNum
        	 int lastrowCount = sourceSheet.getLastRowNum() + 1;
        	 //range "A11:H"+lastrowCount
        	 //CellRangeAddress cellRangeAddressSource = CellRangeAddress.valueOf("A11:H"+lastrowCount);
        	 CellRangeAddress cellRangeAddressSource = CellRangeAddress.valueOf("A2:E"+lastrowCount);
        	//get sheet number 0
        	 XSSFSheet sourceSheet0 =  workbook.getSheetAt(0);
        	//get lastRowNum sheet number 0
        	 int lastrowCountSheet0 = sourceSheet0.getLastRowNum() + 1;
        	 //range "A"+lastrowCountSheet0+1
        	 CellRangeAddress cellRangeAddressDestination = CellRangeAddress.valueOf("A"+(lastrowCountSheet0+1)+":E"+((lastrowCountSheet0+1+lastrowCount)-2));
        	 //copy cellRangeAddressSource and paste in cellRangeAddressDestination
        	 XSSFRangeCopier xssfRangeCopier = new XSSFRangeCopier(sourceSheet, sourceSheet0);
        	 xssfRangeCopier.copyRange(cellRangeAddressSource, cellRangeAddressDestination, true, true);
        	 
        	 for(int j=1;j<lastrowCount;j++)
        	 {
        		 sourceSheet.removeRow(sourceSheet.getRow(j));
        	 }
        	 
        }
        
        File file = new File("C:Users/csra/Desktop/test1.xlsx");
        file.getParentFile().mkdirs();

        FileOutputStream outFile = new FileOutputStream(file);
        workbook.write(outFile);
        workbook.close();
        
        
        
        
        
        
        
        // Get iterator to all the rows in current sheet
//        Iterator<Row> rowIterator = sheet.iterator();
//
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//            // Get iterator to all cells of current row
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//
//                // Change to getCellType() if using POI 4.x
//                CellType cellType = cell.getCellType();
//
//                switch (cellType) {
//						                case _NONE:
//						                    System.out.print("");
//						                    System.out.print("\t");
//						                    break;
//						                case BLANK:
//						                    System.out.print("");
//						                    System.out.print("\t");
//						                    break;
//						                case STRING:
//						                    System.out.print(cell.getStringCellValue());
//						                    System.out.print("\t");
//						                    break;
//						                case NUMERIC:
//						                    System.out.print(cell.getNumericCellValue());
//						                    System.out.print("\t");
//						                    break;
//						                case ERROR:
//						                    System.out.print("!");
//						                    System.out.print("\t");
//						                    break;
//						                }
//
//	            }
//	            System.out.println("");
//        }

	}
	
	

}
