package excelOperations;

public class WriteOperations {

	/**
	 * use this method to add row's into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 */
	public void writeInToExcel(String filePath, String sheetName, Object[][] dataToWrite) {

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

	/**
	 * use this method to update the particular Cell value
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */
	public void updateCellValue(String filePath, String sheetName, String data, int rowIndex,
			int colIndex) {

		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get the Cell number using getRow and getCell */

		/* Step - 6 : Update the cell */

		/* Step - 7 : Close input stream */

		/* Step - 8 : Creating output stream and writing the updated workbook */

		/* Step - 9 : Close the workbook and output stream */
	}

	public void addColumn(String filePath, String sheetName, String[] colValues) {
		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get all the rows and add a new cell to it at the end */

		/* Step - 6 : Close input stream */

		/* Step - 7 : Creating output stream and writing the updated workbook */

		/* Step - 8 : Close the workbook and output stream */

	}

	public void run() {
		// Call the desired methods
	}
}
