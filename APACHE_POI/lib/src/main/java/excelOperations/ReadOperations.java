package excelOperations;

import java.io.IOException;

public class ReadOperations {

	/**
	 * use this method to read the complete excel file
	 * 
	 * @param filePath
	 * @param sheetName
	 * @throws IOException
	 */
	public void readCompleteExcel(String filePath, String sheetName) throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */


		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */


		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		/* Step - 4 : Get the last row number */

		/*
		 * Step - 5 : Get the last cell number
		 */

		/*
		 * Step - 6 : Use a for each loop to iterate the row a. get the row b. using for each loop
		 * iterate over Cell of the row c. using switch statement check Cell type d. print the cell
		 * value
		 *
		 */
	}

	/**
	 * use this method to read the row values from excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @throws IOException
	 */
	public void getRowValue(String filePath, String sheetName, int rowIndex) throws IOException {
		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */


		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		/* Step - 4 : Get the desire row */

		/*
		 * Step - 5 : Iterate over over each Cell using for each loop a. using switch statement
		 * check Cell type b. print the cell value
		 */
	}

	/**
	 * use this method to read column value
	 *
	 * @param filePath
	 * @param sheetName
	 * @param columnIndex
	 * @throws IOException
	 */
	public void getColunmValue(String filePath, String sheetName, int columnIndex)
			throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */


		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		/*
		 * Step - 4 : Using for each loop iterate over each row a. using for each loop iterate over
		 * each Cell of row b. Compare the column index for which you want print the values. if,
		 * match found print the cell value
		 */
	}

	/**
	 * use this method to read a particular Cell value
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */
	public void getCellValue(String filePath, String sheetName, int rowIndex, int colIndex)
			throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		/*
		 * Step - 4 : Get the row from which you want to read the cell data
		 */

		/* Step - 5 : Get the Cell value by passing the column index */

		/* Step - 6 : Print the cell value */
	}

	public void run() {
		// Call the desired methods
	}
}
