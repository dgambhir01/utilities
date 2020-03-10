package Util.util_helper;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProviderDemo {

	private String filePath = "./DataProvider.xlsx";
	private String sheetName = "demo";

	@DataProvider(name = "excelData")
	public Object[][] readExcel() throws InvalidFormatException, IOException {
		return ReadExcel.readExcel(filePath, sheetName);
	}

	// Test method
	@Test(dataProvider = "excelData")

	// Here are my all parameters from excel sheet:
	public void useInScript(String userNAME, String Password) throws Exception {

		System.out.println("Username is:> " + userNAME);
		System.out.println("Password is:> " + Password);

	}

}
