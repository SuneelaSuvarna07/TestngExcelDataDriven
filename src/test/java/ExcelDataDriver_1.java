import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataDriver_1 {

	@Test(dataProvider = "ExcelData")
	public void test(String a, String b, String c) {
		System.out.println(a + " " + b + " " + c);

	}

	@DataProvider(name = "ExcelData")
	public Object[][] getExcelData() throws IOException {
		DataFormatter formatter = new DataFormatter();
		FileInputStream fis = new FileInputStream(
				System.getProperty("user.dir") + File.separator + "data" + File.separator + "ExcelAutomation.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		Object[][] data = new Object[rowCount - 1][colCount];

		for (int i = 0; i < rowCount - 1; i++) {
			row = sheet.getRow(i + 1);
			for (int j = 0; j < colCount; j++) {
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}

		}
		return data;
	}

}
