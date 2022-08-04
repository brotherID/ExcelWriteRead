//package org.manage.file.ManageFile;
//
//import com.aspose.cells.Cells;
//import com.aspose.cells.Range;
//import com.aspose.cells.Workbook;
//import com.aspose.cells.Worksheet;
//
//public class CombineMultipleWorksheets {
//	public static void main(String[] args) throws Exception {
//
//		// The path to the documents directory.
//
//		Workbook workbook = new Workbook("C:/sara/capture-web/Cheque/Cheque1.xlsx");
//
//		Workbook destWorkbook = new Workbook();
//
//		Worksheet destSheet = destWorkbook.getWorksheets().get(0);
//	
//
//		int totalRowCount = 0;
//
//		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
//			Worksheet sourceSheet = workbook.getWorksheets().get(i);
//
//			Cells cells = sourceSheet.getCells();
//
//			if (i != 0) {
//				cells.deleteRow(0);
//			}
//
//			Range sourceRange = cells.getMaxDisplayRange();
//
//			Range destRange = destSheet.getCells().createRange(sourceRange.getFirstRow() + totalRowCount,
//					sourceRange.getFirstColumn(), sourceRange.getRowCount(), sourceRange.getColumnCount());
//
//			destRange.copy(sourceRange);
//
//			totalRowCount = sourceRange.getRowCount() + totalRowCount;
//		}
//
//		destWorkbook.save("C:/sara/capture-web/Cheque/Cheque-output.xlsx");
//
//	}
//
//}