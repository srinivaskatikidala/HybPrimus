package Utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {
	// global variables
	Workbook wb;

	// it is a constructor.
	public ExcelFileUtil() throws Throwable {
		FileInputStream fi = new FileInputStream("F:\\Project\\Automation\\Stock_AccountingMVN\\TestInput\\InputSheet.xlsx");
		wb = new XSSFWorkbook(fi);

	}

	// for row count
	public int rowCount(String sheetname) {

		return wb.getSheet(sheetname).getLastRowNum();

	}

	// for column count
	public int colCount(String sheetname) {

		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}

	// to handle string and numeric data
	public String getData(String sheetname, int row, int column) {

		String data= "";

		if (wb.getSheet(sheetname).getRow(row).getCell(column).getCellType() == Cell.CELL_TYPE_NUMERIC) {

			int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();

			data = String.valueOf(celldata);

		} else {

			data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();

		}

		return data;

	}

	// for writing data with color in sheet

	public void SetData(String sheetname, int row, int column, String status) throws Exception {

		Sheet sh = wb.getSheet(sheetname);
		Row rownum = sh.getRow(row);
		Cell cell = rownum.createCell(column);
		cell.setCellValue(status);

		if (status.equalsIgnoreCase("pass")) {
			// create cell style
			CellStyle style = wb.createCellStyle();
			// create a font
			Font font = wb.createFont();
			// apply color to text
			font.setColor(IndexedColors.GREEN.getIndex());
			// apply bold font style, default bold statement is 'false'
			font.setBold(true);
			// set font
			style.setFont(font);
			// set style
			rownum.getCell(column).setCellStyle(style);

		} else if (status.equalsIgnoreCase("fail")) {

			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);

		}

		else if (status.equalsIgnoreCase("not executed")) {

			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);

		}

		FileOutputStream fos = new FileOutputStream("F:\\Project\\Automation\\Stock_AccountingMVN\\TestOutput.xlsx");
	
		wb.write(fos);
		fos.close();
	

	}

}


