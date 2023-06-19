package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormattingOperations {

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 * @throws FileNotFoundException
	 */

	public void applyColorToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			IndexedColors fillColor, FillPatternType fillPatternType) throws Exception {
		/* Step - 1 : Creating file object of existing excel file */
		File file = new File(filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream fis = new FileInputStream(file);

		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);

		/* Step - 4 : Reading first sheet of excel file */
		XSSFSheet xSheet = xWorkbook.getSheet(sheetName);

		/* Step - 5 : Get the Cell number using getRow and getCell */
		int rowCount = xSheet.getLastRowNum();
		int columnCount = xSheet.getRow(0).getLastCellNum();
		XSSFCell xRow = xSheet.getRow(rowIndex).getCell(colIndex);

		/* Step - 6 : Create the cell style sheet */
		XSSFCellStyle xCellStyle=xWorkbook.createCellStyle();

		/* Step - 7 : Set background color */
		xCellStyle.setFillBackgroundColor(fillColor.getIndex());

		/* Step - 8 : Set fill pattern */
		xCellStyle.setFillPattern(fillPatternType);

		/* Step - 9 : Apply the style to Cell */
		
		xRow.setCellStyle(xCellStyle);
		/* Step - 10 : Close input stream */
		fis.close();

		/* Step - 11 : Creating output stream and writing the updated workbook */
		FileOutputStream fos = new FileOutputStream(file);
		xWorkbook.write(fos);
		xWorkbook.close();
		fos.close();

		/* Step - 12 : Close the workbook and output stream */

	}

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */
	public void applyAlignmentToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			HorizontalAlignment horizontalAlignment) {

		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get the Cell number using getRow and getCell */

		/* Step - 6 : Create the cell style sheet */

		/* Step - 7 : Set Alignment */

		/* Step - 8 : Apply the style to Cell */

		/* Step - 9 : Close input stream */

		/* Step - 10 : Creating output stream and writing the updated workbook */

		/* Step - 11 : Close the workbook and output stream */

	}

	/**
	 * use this method to add row's and apply font color into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 */
	public void applyFontToRow(String filePath, String sheetName, Object[][] dataToWrite,
			String fontName) {

		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Getting the last row number of existing records */

		/**
		 * Step - 6 : Iterating dataToWrite to update* a.Create new row from the next row count
		 * b.Creating new cell and setting the value
		 */

		/* Step - 7 : Close input stream */

		/* Step - 8 : Create output stream and writing the updated workbook */

		/* Step - 9 : Close the workbook and output stream */
	}

	public void validateFormattingUpdates(String filePath, String sheetName) {
		// Creating file object of existing excel file
		File fileName = new File(filePath);

		try {

			FileInputStream file = new FileInputStream(fileName);

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet(sheetName);

			/* Verify setting color to cell B2 */
			XSSFCell cell2Style = sheet.getRow(1).getCell(1);

			XSSFCellStyle style = cell2Style.getCellStyle();

			System.out.println(
					"Color of cell B2 is: " + style.getFillBackgroundColorColor().getARGBHex());

			/* Verify setting alignment of "Survey" column to right alignment */
			cell2Style = sheet.getRow(0).getCell(3);
			style = cell2Style.getCellStyle();

			System.out.println(
					"Alignment of cell A4 (Survey) is: " + style.getAlignment().toString());

			// Close input stream
			file.close();

			// Crating output stream and writing the updated workbook
			FileOutputStream os = new FileOutputStream(fileName);
			workbook.write(os);

			// Close the workbook and output stream
			workbook.close();
			os.close();

			System.out.println("Excel file has been updated successfully.");

		} catch (Exception e) {
			System.err.println("Exception while updating an existing excel file.");
			e.printStackTrace();
		}
	}
	
	public void run() throws Exception {
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";
		// Call the desired methods
      this.applyColorToCell(filePath, worksheetName, 1, 1, IndexedColors.BRIGHT_GREEN,
	  FillPatternType.THICK_BACKWARD_DIAG);
		/* Add your logic to make the formatting updates above this line */

		// Utility method to verify formatting updates
		// this.validateFormattingUpdates(filePath, worksheetName);
	}
}
