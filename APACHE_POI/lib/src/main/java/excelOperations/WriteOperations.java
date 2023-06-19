package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperations {

	/**
	 * use this method to add row's into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 * @throws FileNotFoundException
	 */
	public void writeInToExcel(String filePath, String sheetName, Object[][] dataToWrite) throws Exception {

		/* Step - 1 : Creating file object of existing excel file */
		File file = new File (filePath);
		

		/* Step - 2 : Creating input stream */
        FileInputStream fis = new FileInputStream(file);
		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);

		/* Step - 4 : Reading first sheet of excel file */
		XSSFSheet xSheet = xWorkbook.getSheet(sheetName);

		/* Step - 5 : Getting the last row number of existing records */
		int rowCount=xSheet.getLastRowNum();

		/**
		 * Step - 6 : Iterating dataToWrite to update* a.Create new row from the next row count
		 * b.Creating new cell and setting the value
		 */
		for (Object[] o : dataToWrite){
			XSSFRow xRow=xSheet.createRow(++rowCount);
			int columnIndex=0;
			for (Object info :o){
				XSSFCell xCell=xRow.createCell(columnIndex++);
				if (info instanceof String) {
					xCell.setCellValue((String) info);
				} else if (info instanceof Integer) {
					xCell.setCellValue((Integer) info);
				} else if (info instanceof Double) {
					xCell.setCellValue((Double) info);
				}	
			}
		}

		/* Step - 7 : Close input stream */
		fis.close();
 

		/* Step - 8 : Create output stream and writing the updated workbook */
		FileOutputStream fos = new FileOutputStream(file);
		xWorkbook.write(fos);
		xWorkbook.close();
		fos.close();

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
	 * @throws FileNotFoundException
	 */
	public void updateCellValue(String filePath, String sheetName, String data, int rowIndex,
			int colIndex) throws Exception {

		/* Step - 1 : Creating file object of existing excel file */
		File file = new File (filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream fis = new FileInputStream(file);

		/* Step - 3 : Creating workbook from input stream */
        XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);
		/* Step - 4 : Reading first sheet of excel file */
        XSSFSheet xSheet = xWorkbook.getSheet(sheetName);
		/* Step - 5 : Get the Cell number using getRow and getCell */
        XSSFRow xRow=xSheet.getRow(rowIndex);
		XSSFCell Xcell = xRow.getCell(colIndex);
		/* Step - 6 : Update the cell */
		Xcell.setCellValue(data);

		/* Step - 7 : Close input stream */
		fis.close();
		FileOutputStream fos = new FileOutputStream(file);
		xWorkbook.write(fos);
		xWorkbook.close();
		fos.close();

		/* Step - 8 : Creating output stream and writing the updated workbook */

		/* Step - 9 : Close the workbook and output stream */
	}

	public void addColumn(String filePath, String sheetName, String[] colValues) throws Exception {
		/* Step - 1 : Creating file object of existing excel file */
		File file = new File(filePath);
		FileInputStream fis = new FileInputStream(file);

		/* Step - 2 : Creating input stream */
		XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);
		XSSFSheet xSheet = xWorkbook.getSheet(sheetName);
		int rowCount = xSheet.getLastRowNum();
		int colCount= xSheet.getRow(0).getLastCellNum();
        
		// XSSFCell xCell = xSheet.getRow(0).createCell(colCount, CellType.STRING);
		int index=0;
		for(int i =0; i<colValues.length;i++){
			XSSFCell xCell = xSheet.getRow(index).createCell(colCount, CellType.STRING);
			xCell.setCellValue(colValues[i]);
			index++;
			if (index>rowCount){
				break;
			}

		}
		fis.close();
		FileOutputStream fos = new FileOutputStream(file);
		xWorkbook.write(fos);
		xWorkbook.close();
		fos.close();



		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get all the rows and add a new cell to it at the end */

		/* Step - 6 : Close input stream */

		/* Step - 7 : Creating output stream and writing the updated workbook */

		/* Step - 8 : Close the workbook and output stream */

	}

	public void run() {
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";

		// New students records to update in excel file
		Object[][] countryRecord = {{"UK", "London", "6.72", "15-02-2021"},
				{"US", "Washington,D.C", "32.95", "09-02-2021"}};
		String[] colValues =
				{"Area (Km2)", "3287000", "54394", "30688", "302068", "17100000", "42933"};

		// Add given rows into existing worksheet “Country Population”
		try {
			//this.writeInToExcel(filePath, worksheetName, countryRecord);
			//this.updateCellValue(filePath, worksheetName, "Pronoy", 1, 1);
			
		this.addColumn(filePath, worksheetName, colValues);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
