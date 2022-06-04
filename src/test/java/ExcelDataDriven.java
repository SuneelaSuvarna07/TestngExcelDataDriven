import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataDriven {

	@Test(dataProvider = "getData")
	public void testExcel(String greeting, String comm, String id) {

		
		System.out.println(greeting + " | " + comm + " | " + id);

	}

	@DataProvider
	public Object[][] getData() throws IOException {

		DataFormatter formatter = new DataFormatter();
		FileInputStream fis = new FileInputStream(
				System.getProperty("user.dir") + File.separator + "Data" + File.separator + "ExcelAutomation.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println(rowCount);
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		System.out.println(colCount);
		Object data[][] = new Object[rowCount - 1][colCount];
		for (int i = 0; i < rowCount - 1; i++) {
			row = sheet.getRow(i + 1);
			for (int j = 0; j < colCount; j++) {

				System.out.println(row.getCell(j));
				XSSFCell cellValue = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cellValue);

			}
		}
		return data;

	}

}
