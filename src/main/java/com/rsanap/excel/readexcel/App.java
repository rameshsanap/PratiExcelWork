package com.rsanap.excel.readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {

	private static final String ROOTFOLDER = "E:\\prati";
	private static final String EXCELFILE = "phoenix.xlsx";
	private static final String FILESEPARATOR = "\\";

	public static void main(String[] args) {
		readExcel();
	}

	public static void readExcel() {
		try {
			FileInputStream file = new FileInputStream(new File(ROOTFOLDER
					+ FILESEPARATOR + EXCELFILE));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			TreeMap<Integer, String> folderNames = new TreeMap<Integer, String>();
			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();
					String folderPath = ROOTFOLDER;
					// Check the cell type and format accordingly

					String folderName = cell.getStringCellValue();
					folderNames.put(cell.getColumnIndex(), folderName);
					for (Integer index = 0; index < cell.getColumnIndex(); index++) {
						folderNames.get(index);
						folderPath = folderPath + FILESEPARATOR
								+ folderNames.get(index);
					}
					folderPath = folderPath + FILESEPARATOR + folderName;
					System.out.println("Createing folder " + folderPath);
					createDirectory(folderPath);

				}
				System.out.println("");
			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void createDirectory(String filePath) {
		File file = new File(filePath);
		if (!file.exists()) {
			if (file.mkdir()) {
				System.out.println("Directory is created!");
			}
		}
	}
}
