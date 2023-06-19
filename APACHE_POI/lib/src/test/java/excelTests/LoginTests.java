package excelTests;

import excelTests.utils.ExcelUtils;
import java.io.IOException;
import java.lang.reflect.Method;
import java.lang.reflect.Parameter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

public class LoginTests {


    @BeforeTest
    public void setupTestData() throws IOException {
        // Set Test Data Excel and Sheet
        System.out.println("************Setup Test Level Data**********");
        ExcelUtils.setExcelFileSheet("UserCreds.xlsx", "Sheet1");
    }
    @AfterMethod
    public void getParamater(Method m){
        System.out.println("Executed Method Name->"+" "+m.getName());
        Parameter para[]=m.getParameters();
        for (int i=0 ; i<para.length;i++){
            System.out.println(para[i].getName());
        }
    }

    @Test(priority = 0, description = "Invalid Login Scenario with wrong username and password.")
    @Parameters ({"para1", "para2"})
    public void invalidUserNameInvalidPassword(@Optional String str,@Optional String str2) throws IOException {
        String result;
        XSSFRow xRow=ExcelUtils.getRowData(1);
        String UserName = xRow.getCell(0).getStringCellValue();
        String password = xRow.getCell(1).getStringCellValue();
        result= (UserName.equals("admin$123")&& password.equals("admin$123")) ? "Pass1" :"False1$";
           
        System.out.println(ExcelUtils.setCellData("Result", 2, result));

    }

    @Test(priority = 0, description = "valid Login Scenario with correct username and password.")
    public void validUserNameValidPassword() throws IOException {
        String result;
        XSSFRow xRow=ExcelUtils.getRowData(2);
        String UserName = xRow.getCell(0).getStringCellValue();
        String password = xRow.getCell(1).getStringCellValue();
        result= (UserName.equals("admin$123")&& password.equals("admin$123")) ? "Pass1" :"False1$";
           
        System.out.println(ExcelUtils.setCellData("Result", 3, result));

    }

    @AfterTest
    public void tearDown() throws IOException {
        ExcelUtils.closeWorkbook();
    }
}
