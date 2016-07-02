package utils.poi;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelMng {

	private String bookPath;

	private List<ExcelSheet> excelSheetList = new ArrayList<>();

	public static ExcelMng createExcelMng (String excelPath) {
		return createExcelMng(excelPath, null);
	}

	public static ExcelMng createExcelMng (String excelPath, String sheetName) {

		ExcelMng excelMng = new ExcelMng();
		try {

			excelMng.bookPath = excelPath;

			Workbook workBook = new XSSFWorkbook(excelPath);

	        for (Sheet sheet : workBook) {

	        	if (sheetName == null || sheetName.equals(sheet.getSheetName())) {

		        	ExcelSheet excelSheet = new ExcelSheet(sheet.getSheetName());

		        	for (Row row : sheet) {

	                	ExcelRow excelRow = new ExcelRow();

		                for (int index = 0; index < row.getLastCellNum(); index++) {

		                	Cell cell = row.getCell(index);

		                	excelRow.addColumnValue(index, excelRow.convertColumnValue(cell));
		                }

		                excelSheet.addExcelRow(excelRow);
		            }

			        excelMng.addExcelSheet(excelSheet);
	        	}
	        }

	        workBook.close();

		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return excelMng;
	}


	public void addExcelSheet (ExcelSheet sheet) {
		excelSheetList.add(sheet);
	}
	@Override
	public String toString() {
		return bookPath + System.lineSeparator() + excelSheetList.toString();
	}

}
