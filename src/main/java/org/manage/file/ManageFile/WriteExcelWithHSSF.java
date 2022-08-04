package org.manage.file.ManageFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.manage.file.ManageFile.dao.ChequeDao;
import org.manage.file.ManageFile.model.Cheque;

import lombok.extern.log4j.Log4j2;

@Log4j2
public class WriteExcelWithHSSF 
{
    
	private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }
	
	
    public static void main( String[] args ) throws IOException
    {
    	
    	HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Cheque sheet");

        List<Cheque> list = ChequeDao.listCheques();

        int rownum = 0;
        Cell cell;
        Row row;
        //
        HSSFCellStyle style = createStyleForTitle(workbook);

        row = sheet.createRow(rownum);

        // identifantCheque
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("idCheque");
        cell.setCellStyle(style);
        // cmc7
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("cmc7");
        cell.setCellStyle(style);
        // endos
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("endos");
        cell.setCellStyle(style);
        // montant
        cell = row.createCell(3, CellType.NUMERIC);
        cell.setCellValue("montant");
        cell.setCellStyle(style);
       
        // Data
        for (Cheque c : list) {
            rownum++;
            row = sheet.createRow(rownum);

            // IdentifantCheque
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(c.getIdentifantCheque());
            // Cmc7
            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(c.getCmc7());
            // Endos
            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(c.getEndos());
            // Montant
            cell = row.createCell(3, CellType.NUMERIC);
            cell.setCellValue(c.getMontant());
        }
        File file = new File("C:/sara/capture-web/Cheque/Cheque.xls");
        file.getParentFile().mkdirs();

        FileOutputStream outFile = new FileOutputStream(file);
        workbook.write(outFile);
        System.out.println("Created file: " + file.getAbsolutePath());
    }
}
