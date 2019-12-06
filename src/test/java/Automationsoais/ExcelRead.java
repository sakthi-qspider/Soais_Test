package Automationsoais;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelRead {
	public String[][] getCellData(String path, String sheetName) {
		FileInputStream stream = new FileInputStream(path);
		Workbook workbook = WorkbookFactory.create(stream);
		Sheet s = workbook.getSheet(sheetName);
		int rowcount = s.getLastRowNum();
		int cellcount = s.getRow(0).getLastCellNum();
		String data[][] = new String[rowcount][cellcount];
		for (int i = 1; i <= rowcount; i++) {
		Row r = s.getRow(i);
		for (int j = 0; j<cellcount; j++) {
		Cell c = r.getCell(j);
		try {
		if (c.) {
		data[i - 1][j] = c.getStringCellValue();
		} 
		else
		{
		data[i - 1][j] = String.valueOf(c.getNumericCellValue());
		}
		} catch (Exception e) {
		e.printStackTrace();
		}
		}
		}
		return data;
		}

	
	
	
}
