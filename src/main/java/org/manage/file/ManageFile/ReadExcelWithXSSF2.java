package org.manage.file.ManageFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelWithXSSF2 {

	public static void main(String[] args) throws IOException {
		// Read XSL file
		FileInputStream inputStreamWorkbookInput = new FileInputStream(
				new File("C:/sara/capture-web/Cheque/Cheque1.xlsx"));

		// Get the workbook instance for XLS file
		XSSFWorkbook workbookInput = new XSSFWorkbook(inputStreamWorkbookInput);

		FileInputStream inputStreamWorkbookOutput = new FileInputStream(
				new File("C:/sara/capture-web/Cheque/Cheque2.xlsx"));

		// Get the workbook instance for XLS file
		XSSFWorkbook workbookOutput = new XSSFWorkbook(inputStreamWorkbookOutput);
		XSSFSheet sheetOutput = workbookOutput.createSheet();

		for (int sheetIndex = 0; sheetIndex < workbookInput.getNumberOfSheets(); sheetIndex++) {
			XSSFSheet sheetInput = workbookInput.getSheetAt(sheetIndex);

//			sheetOutput.copyRows(getAllRowsfromSheet(sheetInput, sheetIndex), 0,
//					false);

		}

		// Get first sheet from the workbook
		XSSFSheet sheet = workbookInput.getSheetAt(0);

		// Get iterator to all the rows in current sheet
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			// Get iterator to all cells of current row
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				// Change to getCellType() if using POI 4.x
				CellType cellType = cell.getCellType();

				switch (cellType) {
				case _NONE:
					System.out.print("");
					System.out.print("\t");
					break;
				case BLANK:
					System.out.print("");
					System.out.print("\t");
					break;
				case STRING:
					System.out.print(cell.getStringCellValue());
					System.out.print("\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					System.out.print("\t");
					break;
				case ERROR:
					System.out.print("!");
					System.out.print("\t");
					break;
				}

			}
			System.out.println("");
		}

	}

	private static List<XSSFRow> getAllRowsfromSheet(XSSFSheet sheetInput, int sheetindex) {
		List<XSSFRow> allRowsfromSheet = new ArrayList<XSSFRow>();
		for (int i = (sheetindex == 0 ? 0 : 1); i <= sheetInput.getPhysicalNumberOfRows(); i++) {
			allRowsfromSheet.add(sheetInput.getRow(i));
		}
		return allRowsfromSheet;
	}

}
