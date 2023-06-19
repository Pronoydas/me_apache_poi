package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadOperations {

	/**
	 * use this method to read the complete excel file
	 * 
	 * @param filePath
	 * @param sheetName
	 * @throws IOException
	 */
	public void readCompleteExcel(String filePath, String sheetName) throws IOException {
         File file = new File(filePath);
		 FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook xworkbook = new XSSFWorkbook(fis);
		XSSFSheet xsheet=xworkbook.getSheet(sheetName);
         
		/* Step - 4 : Get the last row number */
		int rowCount = xsheet.getLastRowNum();
		int columnCount= xsheet.getRow(0).getLastCellNum();
		System.out.println("Row Count ->"+rowCount+" column count->"+columnCount);
        for (int i =0 ; i<=rowCount; i++){
			XSSFRow xrRow=xsheet.getRow(i);
			for (int j =0; j<columnCount;j++){
				XSSFCell xCell = xrRow.getCell(j);
				switch(xCell.getCellType()){
					case STRING :System.out.print(xCell.getStringCellValue());break;
					case NUMERIC : System.out.print(xCell.getNumericCellValue());break;
					case BOOLEAN : System.out.print(xCell.getBooleanCellValue());break ;
					default: System.out.println("Incorrect Data Type");
				}
				System.out.print("|");

			}
			System.out.println();
		}




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
		File file = new File (filePath);
		FileInputStream fis = new FileInputStream(file);

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
	    XSSFWorkbook xworkbook = new XSSFWorkbook(fis);
		XSSFSheet xsheet = xworkbook.getSheet(sheetName);
    
		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 * 
		 */
		XSSFRow xrRow = xsheet.getRow(rowIndex);
		
		int cellCount = xrRow.getLastCellNum();
		
        for(int i=0 ;i<cellCount; i++){
			XSSFCell xCell=xrRow.getCell(i);
			switch(xCell.getCellType()){
				case STRING : System.out.print(xCell.getStringCellValue()+" |");break;
				case BOOLEAN : System.out.print(xCell.getBooleanCellValue()+" |");break;
				case NUMERIC : System.out.print(xCell.getNumericCellValue()+" |");break;
				default : System.out.println("::Data type mismatch::");
			}
			
			
		}
       System.out.println();

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
		File file = new File(filePath);
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);
		XSSFSheet xSheet = xWorkbook.getSheet(sheetName);
		int rowCount = xSheet.getLastRowNum();
		for(int i =0 ; i<=rowCount;i++){
			XSSFRow xRow=xSheet.getRow(i);
			XSSFCell xCell=xRow.getCell(columnIndex);
			switch(xCell.getCellType()){
				case STRING : System.out.println(xCell.getStringCellValue());break;
				case NUMERIC : System.out.println(xCell.getNumericCellValue());break;
			}
		}
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
				File file = new File(filePath);
				FileInputStream fis = new FileInputStream(file);
				XSSFWorkbook xWorkbook = new XSSFWorkbook(fis);
				XSSFSheet xSheet = xWorkbook.getSheet(sheetName);
				XSSFRow xRow=xSheet.getRow(rowIndex);
				XSSFCell xCell=xRow.getCell(colIndex);
				switch(xCell.getCellType()){
					case STRING : System.out.println(xCell.getStringCellValue());break;
					case BOOLEAN : System.out.println(xCell.getBooleanCellValue());break;
					case NUMERIC : System.out.println(xCell.getNumericCellValue());break;
				}
	
	}

	public void run() {
		// Call the desired methods
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";
		try {
			System.out.println("from run");
			this.readCompleteExcel(filePath, worksheetName);
        //   System.out.println("=============== Reading A Row Value =================");
		//   this.getRowValue(filePath, worksheetName, 1);
		//   System.out.println("=============== Reading A Column Value =================");
        //   this.getColunmValue(filePath,worksheetName,0);
		//   System.out.println("=============== Reading A Cell Value =================");
		//   this.getCellValue(filePath,worksheetName, 1, 1);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
