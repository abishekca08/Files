package jp.co;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {
	public static void main(String[] args) throws Exception {
		// Create blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a blank sheet
		XSSFSheet spreadsheet = workbook.createSheet(" Bug Count ");
		// Create row object
		XSSFRow row;
		// This data needs to be written (Object[])
		Map<String, String> buginfo = new TreeMap<String, String>();
		buginfo.put("32113", "244");
		buginfo.put("32115", "221");
		buginfo.put("34115", "54");
		buginfo.put("43004", "678");
		buginfo.put("89082", "231");
		buginfo.put("32123", "244");
		buginfo.put("32105", "261");
		buginfo.put("32115", "540");
		buginfo.put("41004", "678");
		buginfo.put("83082", "231");
		buginfo.put("35113", "244");
		buginfo.put("37115", "221");
		buginfo.put("38115", "541");
		buginfo.put("43404", "678");
		buginfo.put("89382", "231");
		buginfo.put("32117", "244");
		buginfo.put("32005", "221");
		buginfo.put("34915", "444");
		buginfo.put("43204", "608");
		buginfo.put("79082", "231");
		// Iterate over data and write to sheet
		Set<String> keyid = buginfo.keySet();
		int rowid = 1;
		row = spreadsheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Ticket ID");
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Bug Count");
		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			//Integer objectArr = buginfo.get(key);
			int cellid = 0;
			cell = row.createCell(cellid++);
			cell.setCellValue(key);
			cell1 = row.createCell(cellid++);
			cell1.setCellValue(buginfo.get(key));
		}
		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File("D:\\Files\\CJK\\Task\\BugSheet.xlsx"));
		workbook.write(out);
		out.close();
		System.out.println("Excel written successfully");
	}
}
