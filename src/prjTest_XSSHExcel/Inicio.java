package prjTest_XSSHExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Inicio {

	public static void main(String[] args) {

		Date date = new Date();
		String newString = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
		System.out.print(newString + "\n");
		Inicio p = new Inicio();
		p.pintar();
		date = new Date();
		newString = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
		System.out.print(newString + "\n");
		System.out.print("FIN.");
		
		
	}

	private void pintar() {
		
		int totalFilas = 20000;
		int totalColumnas = 24;
		FileOutputStream fos = null;
		
		try {
			FileInputStream inputStream = new FileInputStream("mytemplate.xlsx");
        	XSSFWorkbook wb_template = new XSSFWorkbook(inputStream);
        	inputStream.close();

        	SXSSFWorkbook workbook = new SXSSFWorkbook(wb_template); 
        	workbook.setCompressTempFiles(true);

        	SXSSFSheet sheet = (SXSSFSheet) workbook.getSheetAt(0);
        	sheet.setRandomAccessWindowSize(100);// keep 100 rows in memory, exceeding rows will be flushed to disk
        
        	Row row;
        	Cell cell;
        	CellStyle style1 = workbook.createCellStyle();
			CellStyle style2 = workbook.createCellStyle();
			CellStyle style3 = workbook.createCellStyle();
			
			establecerStilo1(workbook, style1);
			establecerStilo2(workbook, style2);
			establecerStilo3(workbook, style3);
	
			for( int i= 1; i < totalFilas; i++) {
				row = sheet.createRow(i);
				
				for (int j = 0; j < totalColumnas; j++) {
					cell  = row.createCell(j);
					cell.setCellValue(new HSSFRichTextString("sandeep" + i + "-" + j));
					if (j == 0) {
						cell.setCellStyle(style1);
					} else if (j == 1) {
						cell.setCellStyle(style2);
					} else if (j == 3) {
						cell.setCellStyle(style3);
					}
				}
			}
		   
			fos = new FileOutputStream(new File("myExcelWorkBook.xls"));
			workbook.write(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fos.flush();
				fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void establecerStilo1(SXSSFWorkbook workbook, CellStyle style) {
		DataFormat df = workbook.createDataFormat(); 
		style.setAlignment((short)1);
		style.setBorderBottom((short)1);
		style.setBorderLeft((short)1);
		style.setBorderRight((short)1);
		style.setBorderTop((short)1);
		style.setDataFormat(df.getFormat("0.0"));
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		style.setFillBackgroundColor((short) 21);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
	}
	public void establecerStilo2(SXSSFWorkbook workbook, CellStyle style) {
		DataFormat df = workbook.createDataFormat(); 
		style.setAlignment((short)1);
		style.setBorderBottom((short)1);
		style.setBorderLeft((short)1);
		style.setBorderRight((short)1);
		style.setBorderTop((short)1);
		style.setDataFormat(df.getFormat("0,00"));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillBackgroundColor((short) 11);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
	}
	public void establecerStilo3(SXSSFWorkbook workbook, CellStyle style) {
		DataFormat df = workbook.createDataFormat(); 
		style.setAlignment((short)1);
		style.setBorderBottom((short)1);
		style.setBorderLeft((short)1);
		style.setBorderRight((short)1);
		style.setBorderTop((short)1);
		style.setDataFormat(df.getFormat("0.0"));
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		style.setFillBackgroundColor((short) 33);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
	}
	public void establecerStilo4(SXSSFWorkbook workbook, CellStyle style) {
		style.setAlignment((short)1);
		style.setBorderBottom((short)1);
		style.setBorderLeft((short)1);
		style.setBorderRight((short)1);
		style.setBorderTop((short)1);
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		style.setFillBackgroundColor((short) 2);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
	}
	
}
