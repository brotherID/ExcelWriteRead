package org.manage.file.ManageFile;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelRemoveMergedRegion {
	
	static boolean removeMergedRegion(Sheet sheet, CellRangeAddress mergedRegionToRemove) {
		  boolean removed = false;
		  for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
		   CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
		   if (mergedRegionToRemove.equals(mergedRegion)) {
		    sheet.removeMergedRegion(i);
		    removed = true;
		    break;
		   }
		  }
		  return removed;
		 }

}
