package mainPackage;

import java.util.HashMap;

public class Constants {
	// Assigning Constant values

	/*Here you mention the below Constant values
	 
	1. In testCaseFile mention your testcase sheet path
	2. Mention drivers path like IE, Chrome and Firefox
	3. In screenPath mention your report path where you want to save your screenshots
	4. Mention your cloud url, username and password
	
	*/

	public final String testCaseFile = "src\\testdata\\Assignment.xlsx";
	public String testCaseSheetName = "Test Case";
	public String testStepSheetName = "Test Step";
	public String testDataSheetName = "Test Data";
	public String chromeDriver = "webdriver.chrome.driver";
	public String chromeDriverPath = "src\\libs\\chromedriver.exe";
	public String screenPath = "src\\reports\\";
	public String pass = "Pass";
	public String fail = "Fail";
	public static String CLIPath = "C:\\dc\\dc-cli.exe";
	public static String ipAddress = "10.87.240.42";
	public static String userName = "";
	public static String password = "";
	public static String APIKey = "";
	public static HashMap<String, String> hash;

	public static void getURL() {
		hash = new HashMap<String, String>();
		//hash.put("Google_Pixel2", "bc6f2b10-0ff0-4858-8158-b9c4798101cc");
		
		/*
		  
		  Add your device with their name of the device and its ID from Cloud
		  Please refer the above for reference.

		*/
		
	}

}
