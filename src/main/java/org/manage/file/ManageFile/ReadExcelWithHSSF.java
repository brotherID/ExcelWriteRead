package org.manage.file.ManageFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRangeCopier;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ReadExcelWithHSSF {

	public static void main(String[] args) throws IOException {
		// Read XSL file
		FileInputStream inputStream = new FileInputStream(new File("C:Users/csra/Desktop/bordereau.xls"));
		// Get the workbook instance for XLS file
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			HSSFSheet sourceSheet = workbook.getSheetAt(i);
			ExcelRemoveMergedRegion.removeMergedRegion(sourceSheet, CellRangeAddress.valueOf("G3:H3"));
			ExcelRemoveMergedRegion.removeMergedRegion(sourceSheet, CellRangeAddress.valueOf("A8:B8"));
			ExcelRemoveMergedRegion.removeMergedRegion(sourceSheet, CellRangeAddress.valueOf("C8:D8"));
		}
		for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
			// get active sheet
			HSSFSheet sourceSheet = workbook.getSheetAt(i);
			// get lastRowNum
			int lastSourceRowCount = sourceSheet.getLastRowNum()+1;
			// range "A11:H"+lastrowCount
			CellRangeAddress cellRangeAddressSource = CellRangeAddress.valueOf("A11:H" + lastSourceRowCount);
			// get sheet number 0
			HSSFSheet destinationSheet0 = workbook.getSheetAt(0);
			// get lastRowNum sheet number 0
			int lastDestinationRowCountSheet0 = destinationSheet0.getLastRowNum()+1;
			
			// range "A"+lastrowCountSheet0+1
			CellRangeAddress cellRangeAddressDestination = CellRangeAddress
					.valueOf("A" + (lastDestinationRowCountSheet0 +(i!=1 ? 1:0)) + ":H"
							+ (lastDestinationRowCountSheet0 + (lastSourceRowCount - 10 -1)));
			
			// copy cellRangeAddressSource and paste in cellRangeAddressDestination
			HSSFRangeCopier hssfRangeCopier = new HSSFRangeCopier(sourceSheet, destinationSheet0);
			hssfRangeCopier.copyRange(cellRangeAddressSource, cellRangeAddressDestination, true, true);
//			for (int j = 10; j < lastSourceRowCount; j++) {
//				sourceSheet.removeRow(sourceSheet.getRow(j));
//			}

		}
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			HSSFSheet sourceSheet = workbook.getSheetAt(i);
			sourceSheet.addMergedRegion(CellRangeAddress.valueOf("A8:B8"));
			sourceSheet.addMergedRegion(CellRangeAddress.valueOf("C8:D8"));
			sourceSheet.addMergedRegion(CellRangeAddress.valueOf("G3:H3"));

		}
		File file = new File("C:Users/csra/Desktop/destination.xls");
		file.getParentFile().mkdirs();
		FileOutputStream outFile = new FileOutputStream(file);
		workbook.write(outFile);
		workbook.close();
	}

	private static List<List<String>> copyData(HSSFSheet sourceSheet) {
		List<List<String>> data = new ArrayList<List<String>>();
		int FirstSourceRowNumber = 10;
		int lastSourceRowCount = sourceSheet.getLastRowNum() + 1;
		for (int i = FirstSourceRowNumber; i < lastSourceRowCount; i++) {
			HSSFRow row = sourceSheet.getRow(i);
			List<String> rowDataList = new ArrayList<String>();
			HSSFCell cell = null;
			int fCellNum = row.getFirstCellNum();
			int lCellNum = row.getLastCellNum();
			for (int j = fCellNum; j < lCellNum; j++) {
				cell = row.getCell(j);
				rowDataList.add(cell.toString());
			}
			data.add(rowDataList);
		}

		return data;
	}

//	    private static void createNewRow(List<List<String>> dataRows,HSSFSheet destinationSheet,int lastDestinationRowCountSheet0) {
//	    	HSSFCell cell=null;
//	    	HSSFRow  row= null;
//	    	List<String> listCells = new ArrayList<String>();
//	    	for(int i=0;i<dataRows.size();i++)
//	    	{
//	    		row = destinationSheet.createRow(lastDestinationRowCountSheet0+1);
//	    		for(int j=0;j<dataRows.get(i).size();j++)
//		    	{
//	    			listCells.add(dataRows.get(i).get(j));
//	    			System.out.println("element "+listCells.get(j));
//	    			String element = listCells.get(j);
//	    			
//            		cell = row.createCell(j);
//            		cell.setCellValue(element);
//	    			
//		    	}
//	    		
//	    	}
//	    	
//	    	
//	    	
//	    	
//	    	
//			for (Integer key : keyset) { 
//	            // this creates a new row in the sheet 
//	            excelsData = dataRows.get(key);
//	            for (List<String> listData:dataRows) {
//	            	row= sheet.createRow(rownum);
//	            	int j=0;
//	            	for(String data:listData) {
//	            		cell = row.createCell(j);
//	            		cell.setCellValue(data);
//	                    cell.setCellStyle(headerCellStyle1);
//	                    cell=null;
//	                    System.gc();
//	                    j++;
//	            	}
//	            	row=null;
//	            	rownum++;
//				}
//	            
//	        }
//	    	
//		}
//				

	private static void copyRow(HSSFWorkbook workbook, HSSFSheet sourceSheet, int sourceRowNum, int destinationRowNum,
			HSSFSheet destinationSheet0) {
		// Get the source / new row
		HSSFRow newRow = destinationSheet0.getRow(destinationRowNum);
		HSSFRow sourceRow = sourceSheet.getRow(sourceRowNum);

		// Loop through source columns to add to new row
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			// Grab a copy of the old/new cell
			HSSFCell oldCell = sourceRow.getCell(i);
			HSSFCell newCell = newRow.createCell(i);

			// Copy style from old cell and apply to new cell
			HSSFCellStyle newCellStyle = workbook.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
			newCell.setCellStyle(newCellStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());
			newCell.setCellValue(oldCell.getRichStringCellValue());
		}
	}
}