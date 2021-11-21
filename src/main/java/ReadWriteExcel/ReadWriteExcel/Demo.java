package ReadWriteExcel.ReadWriteExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		String readerFilePath = "./src/main/java/ReadWriteExcel/ReadWriteExcel/Read.xlsx";
		String writerFilePath = "./src/main/java/ReadWriteExcel/ReadWriteExcel/Write.xlsx";
		String data[][] = readData(readerFilePath, "Sheet1");
		writeData(writerFilePath, "Sheet1", data, readerFilePath, "Sheet1");
		System.out.println("Im done...");

	}

	public static void writeData(String writerFilePath, String writerSheet, String data[][], String readerFilePath,
			String readerSheet) throws IOException {

		FileInputStream fs = new FileInputStream(writerFilePath);

		Workbook wb = new XSSFWorkbook(fs);
		Sheet sheet1 = wb.getSheet(writerSheet);

		int rowCount = getRows(readerFilePath, readerSheet);
		int columnCount = getColms(readerFilePath, readerSheet);

		for (int i = 0; i < rowCount + 1; i++) {
			Row row = sheet1.createRow(i);
			for (int j = 0; j < columnCount; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue(data[i][j]);
			}
		}
		FileOutputStream fos = new FileOutputStream(writerFilePath);
		wb.write(fos);
		fos.close();
	}

	public static String[][] readData(String filePath, String sheetName) throws IOException {
		FileInputStream fs = new FileInputStream(filePath);
		// Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		int rowNos = sheet.getLastRowNum();
		Row row = sheet.getRow(0);
		int column = row.getLastCellNum();

		String data[][] = new String[rowNos + 1][column + 1];

		for (int i = 0; i <= rowNos; i++) {
			for (int j = 0; j < column; j++) {
				// System.out.println(sheet.getRow(i).getCell(j));
				data[i][j] = sheet.getRow(i).getCell(j).getRichStringCellValue().toString();
			}
		}
		return data;
	}

	public static int getRows(String filePath, String sheetName) {
		String readerFilePath = "./src/main/java/com/excel/lib/util/Read.xlsx";

		XSSFSheet sheet = null;
		try {
			FileInputStream fs = new FileInputStream(filePath);
			// Creating a workbook
			XSSFWorkbook workbook = new XSSFWorkbook(fs);
			sheet = workbook.getSheet(sheetName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return sheet.getLastRowNum();
	}

	public static int getColms(String filePath, String sheetName) {

		XSSFSheet sheet = null;
		try {
			FileInputStream fs = new FileInputStream(filePath);
			// Creating a workbook
			XSSFWorkbook workbook = new XSSFWorkbook(fs);
			sheet = workbook.getSheet(sheetName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();

		}
		return sheet.getRow(0).getLastCellNum();
	}
}
