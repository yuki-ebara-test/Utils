package utils.poi;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelRow {

	private Map<String, String> row = new HashMap<>();

	public void addColumnValue (int index, String value) {
		row.put(String.valueOf(index), value);
	}
	public String convertColumnValue(Cell cell) {

		if (cell == null) return "";

		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				return cell.getRichStringCellValue() == null ? "" : String.valueOf(cell.getRichStringCellValue());
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue() == null ? "" : String.valueOf(cell.getDateCellValue());
				} else {
					return String.valueOf(cell.getNumericCellValue());
				}
			case Cell.CELL_TYPE_BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case Cell.CELL_TYPE_FORMULA:
				return cell.getCellFormula() == null ? "" : String.valueOf(cell.getCellFormula());
			default:
				return "";
		}
	}
	@Override
	public String toString() {
		return System.lineSeparator() + "  " + row.toString() ;
	}

}
