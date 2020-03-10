package Util.util_helper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class ReadAllSheets {

	public static final String FILE_PATH = "./Demo.xlsx";

	public static void main(String[] args) throws IOException, InvalidFormatException {

		Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));

		// Checking the number of sheets in the Workbook
		System.out.println("This Workbook has " + workbook.getNumberOfSheets() + " Sheets.\n");

		// Here's a sheet Iterator 
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		System.out.println("Name of the sheets are:");
		while (sheetIterator.hasNext()) {
			Sheet sheet = sheetIterator.next();
			System.out.println("=> " + sheet.getSheetName());
		}

		if (workbook.getNumberOfSheets() > 0) {

			for (int k = 0; k < workbook.getNumberOfSheets(); k++) {

				Sheet sheet = workbook.getSheetAt(k);

				// You will need a DataFormatter in case you have a date to format and convert
				// it's value to a String
				DataFormatter dataFormatter = new DataFormatter();

				System.out.println("\nData from sheet: "+workbook.getSheetAt(k).getSheetName());
				Iterator<Row> rowIterator = sheet.rowIterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					// Now let's iterate over the columns of the current row
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						String cellValue = dataFormatter.formatCellValue(cell);
						System.out.print(cellValue + "\t");
					}
					System.out.println();
				}

				// Closing the workbook
				workbook.close();
			}
		}
	}

}
