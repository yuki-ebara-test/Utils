package test.utils.poi;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Cell;
import org.junit.Test;

import mockit.Mock;
import mockit.MockUp;
import utils.poi.ExcelManager;
import utils.poi.ExcelRow;
import utils.poi.ExcelSheet;

public class ExcelMngTest {

	public static void main (String[] args) {
		try {
			ExcelManager excelManager = ExcelManager.createExcelManager("D:\\pleiades\\workspace\\Utils\\tmp\\aaa.xlsx", "原本");

			System.out.println(excelManager);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void test() {

		try {
			ExcelManager excelManager = ExcelManager.createExcelManager("D:\\pleiades\\workspace\\Utils\\tmp\\POItest.xlsx");

		} catch (Exception e) {
			e.printStackTrace();
			fail();
		}
	}

	@Test
    public void testMock() {
		// 指定クラスの関数の内容を改変
		new MockUp<ExcelSheet>() {
            @Mock
            public void addExcelRow (ExcelRow row) {}
        };
		new MockUp<ExcelRow>() {
            @Mock
            public String convertToString(Cell cell) {
            	return "";
            }
        };

		try {
			ExcelManager target = ExcelManager.createExcelManager("D:\\pleiades\\workspace\\Utils\\tmp\\POItest.xlsx");

			System.out.println(target);

		} catch (Exception e) {
			e.printStackTrace();
			fail();
		}

    }
}
