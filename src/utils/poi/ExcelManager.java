package utils.poi;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelManager {

	private String bookPath;

	private List<ExcelSheet> excelSheetList = new ArrayList<>();

	/**
	 *
	 */
	private ExcelManager () {

	}

	/**
	 *
	 * @param excelPath
	 * @return
	 */
	public static ExcelManager createExcelManager (String excelPath) {
		return createExcelManager(excelPath, null);
	}
	/**
	 *
	 * @param excelPath
	 * @param sheetName
	 * @return
	 */
	public static ExcelManager createExcelManager (String excelPath, String sheetName) {

		if (excelPath == null || excelPath.length() == 0) {
			throw new RuntimeException("パラメータが未設定 excelPath");
		}

		ExcelManager excelMng = new ExcelManager();

		try {

			excelMng.bookPath = excelPath;

			Workbook book = new XSSFWorkbook(excelPath);

	        for (Sheet sheet : book) {

	        	if (sheetName == null || sheetName.equals(sheet.getSheetName())) {

		        	ExcelSheet excelSheet = new ExcelSheet(sheet.getSheetName());

		        	for (Row row : sheet) {

	                	ExcelRow excelRow = new ExcelRow();

		                for (int index = 0; index < row.getLastCellNum(); index++) {

		                	Cell cell = row.getCell(index);

		                	excelRow.addColumnValue(index, excelRow.convertToString(cell));
		                }

		                excelSheet.addExcelRow(excelRow);
		            }

			        excelMng.addExcelSheet(excelSheet);
	        	}
	        }

	        book.close();

		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return excelMng;
	}

	/**
	 *
	 * @param sheet
	 */
	public void addExcelSheet (ExcelSheet sheet) {
		excelSheetList.add(sheet);
	}

	@Override
	public String toString() {
		return bookPath + System.lineSeparator() + excelSheetList.toString();
	}

}
