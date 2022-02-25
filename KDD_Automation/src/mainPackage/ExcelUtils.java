package mainPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class ExcelUtils {
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	FileInputStream fis;
	FileOutputStream fos;
	String TCId, objectName, objectType, action, data, TSId, device, username, password, url, OS, MYCAURL, MYCAUS,
			MYCAPW, description;
	boolean present, runMode, result;
	static int dataRow = 1;
	int cycle = 0;
	int lastRowNo, intData;
	double doubleData;
	public static Document document;
	public static OutputStream file;
	public static Map<String, String> StepWiseResult = new TreeMap<String, String>();
	public static Map<String, String> TestCasePass = new TreeMap<String, String>();
	public static Map<String, String> TestCaseFail = new TreeMap<String, String>();

	// This method is used to read test steps from <testdata>.xls file
	public void readTestcase() throws Exception {
		fis = new FileInputStream(new Constants().testCaseFile);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(new Constants().testCaseSheetName);
		lastRowNo = sheet.getLastRowNum();
		fis.close();
		for (int i = 1; i <= lastRowNo; i++) {
			fis = new FileInputStream(new Constants().testCaseFile);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet(new Constants().testCaseSheetName);
			row = sheet.getRow(i);
			cell = row.getCell(0);
			TCId = cell.getStringCellValue();

			if (row.getCell(2).getStringCellValue().equalsIgnoreCase("yes")) {
				runMode = true;
			} else {
				runMode = false;
			}

			device = row.getCell(4).getStringCellValue();
			cycle = (int) row.getCell(3).getNumericCellValue();
			OS = row.getCell(5).getStringCellValue();
			fis.close();

			if (runMode) {
				System.out.print(TCId + "Executing...\n");
				file = new FileOutputStream(new File(new Constants().screenPath + TCId + ".pdf"));
				document = new Document();
				PdfWriter.getInstance(document, file);
				document.open();

				for (int k = 0; k < cycle; k++) {
					fis = new FileInputStream(new Constants().testCaseFile);
					workbook = new XSSFWorkbook(fis);
					sheet = workbook.getSheet(new Constants().testDataSheetName);
					row = sheet.getRow(dataRow++);

					if (row.getCell(0).getStringCellValue().equalsIgnoreCase(TCId)) {
						try {
							url = row.getCell(3).getStringCellValue().trim();
							username = row.getCell(1).getStringCellValue().trim();
							password = row.getCell(2).getStringCellValue().trim();
						} catch (Exception e) {
							username = "";
							password = "";
						}
					}

					fis.close();

					fis = new FileInputStream(new Constants().testCaseFile);
					workbook = new XSSFWorkbook(fis);
					sheet = workbook.getSheet(new Constants().testStepSheetName);

					for (int j = 1; j <= sheet.getLastRowNum(); j++) {
						row = sheet.getRow(j);
						cell = row.getCell(0);
						if (TCId.equalsIgnoreCase(cell.getStringCellValue())) {
							runTestStep(row);
						}
					}
					fis.close();
					document.close();
					file.close();
				}
			} else if (!runMode) {
				dataRow += cycle;
			}
		}
	}

	// This method is used to run the test steps
	public void runTestStep(XSSFRow steprow) throws Exception {
		TSId = steprow.getCell(1).getStringCellValue();
		System.out.print(TSId + " Executing...\n");

		try {
			objectName = steprow.getCell(3).getStringCellValue();
		} catch (Exception e) {
			objectName = "";
		}

		try {
			objectType = steprow.getCell(4).getStringCellValue();
		} catch (Exception e) {
			objectType = "";
		}

		try {
			description = steprow.getCell(2).getStringCellValue();
		} catch (Exception e) {
			description = "";
		}

		try {
			action = steprow.getCell(5).getStringCellValue();
		} catch (Exception e) {
			objectType = "";
		}

		try {
			if (action.equalsIgnoreCase("wait") || action.equalsIgnoreCase("scrollDownScreenshot")
					|| action.equalsIgnoreCase("scrollScreenshot") || action.equalsIgnoreCase("swipeUp")
					|| action.equalsIgnoreCase("swipeDown") || action.equalsIgnoreCase("swipeRightLeft")
					|| action.equalsIgnoreCase("swipeLeftRight") || action.equalsIgnoreCase("select")
					|| action.equalsIgnoreCase("splitInteger") || action.equalsIgnoreCase("selectFrameIndex")
					|| action.equalsIgnoreCase("splitString") || action.equalsIgnoreCase("swipeLeftiOS"))
				intData = (int) steprow.getCell(6).getNumericCellValue();
			else if (action.equalsIgnoreCase("input") || action.equalsIgnoreCase("selectByValue")
					|| action.equalsIgnoreCase("pickerfield")) {
				if (steprow.getCell(6).getCellType() == Cell.CELL_TYPE_NUMERIC)
					data = BigDecimal.valueOf(steprow.getCell(6).getNumericCellValue()).toBigInteger().toString();
				else
					data = steprow.getCell(6).getStringCellValue();
			} else
				data = steprow.getCell(6).getStringCellValue();

			if (data.equalsIgnoreCase("username"))
				data = username;
			else if (data.equalsIgnoreCase("password"))
				data = password;
			else if (data.equalsIgnoreCase("url"))
				data = url;
		}

		catch (Exception e) {
			data = "";
		}

		try {
			// Below code is used to perform launching specified device
			if (action.equalsIgnoreCase("open")) {
				if (!data.equalsIgnoreCase("")) {
					new Desktop_ActionKeywords().open(device, url);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("URL not provided...");
				}
			}

			 

			

			else if (action.equalsIgnoreCase("comment")) {
				if (!data.equalsIgnoreCase("")) {
					ExcelUtils.document.add(new Paragraph(" "));
					ExcelUtils.document.add(new Paragraph(data));
					ExcelUtils.document.add(new Paragraph(" "));
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Comments not provided...");
				}
			}   else if (action.equalsIgnoreCase("input")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")
							&& !data.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().input(objectName, objectType, data);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("No input parameters...");
					}
				}

				 else {
					System.out.println("Invalid Browser/OS type");
				}
			}

			else if (action.equalsIgnoreCase("click")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().click(objectName, objectType);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("No input parameters...");
					}
				}

				else {
					System.out.println("Invalid action type");
				}
			} 
			else if (action.equalsIgnoreCase("swipeUp")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} else if (action.equalsIgnoreCase("swipeDown")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} else if (action.equalsIgnoreCase("swipeRightLeft")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} else if (action.equalsIgnoreCase("swipeLeftRight")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} 

			else if (action.equalsIgnoreCase("capture")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} else if (action.equalsIgnoreCase("scrollDownScreenshot_APP")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} else if (action.equalsIgnoreCase("scrollScreenshot")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().CaptureScreen(TCId);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			} 

			// Coming out from frame
			else if (action.equalsIgnoreCase("closeFrame")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().closeFrame();
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} 
			}
			// Perform validateObject action
			else if (action.equalsIgnoreCase("validateObject")) {
				if (OS.equalsIgnoreCase("Browser")) {
					// System.out.println("'" + data + "'");
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						result = new Desktop_ActionKeywords().validateObject(objectName, objectType, TSId);

						if (result) {
							updateStatus(TCId, TSId, steprow, new Constants().pass);
						} else {
							updateStatus(TCId, TSId, steprow, new Constants().fail);
						}
					}
				} 
				else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Validation failed...");
				}
			}

			// Perform validateObjectEnable action
			else if (action.equalsIgnoreCase("ObjectEnable")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						result = new Desktop_ActionKeywords().ObjectEnable(objectName, objectType);
						if (result) {
							updateStatus(TCId, TSId, steprow, new Constants().pass);
						} else {
							updateStatus(TCId, TSId, steprow, new Constants().fail);
						}
					}
				} 
				 else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Validation failed...");
				}
			}

			// Perform validateObjectDisable action
			else if (action.equalsIgnoreCase("ObjectDisable")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						result = new Desktop_ActionKeywords().ObjectDisable(objectName, objectType);
						if (result) {
							updateStatus(TCId, TSId, steprow, new Constants().pass);
						} else {
							updateStatus(TCId, TSId, steprow, new Constants().fail);
						}
					}
				} else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Validation failed...");
				}
			}

			// Perform Validate text
			else if (action.equalsIgnoreCase("validateText")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")
							&& !data.equalsIgnoreCase("")) {
						result = new Desktop_ActionKeywords().validateText(objectName, objectType, data);

						if (result) {
							updateStatus(TCId, TSId, steprow, new Constants().pass);
						} else {
							updateStatus(TCId, TSId, steprow, new Constants().fail);
						}

					}
				}  else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Validation failed...");
				}
			}

			else if (action.equalsIgnoreCase("newTab")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (objectName.equalsIgnoreCase("") && objectType.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().newTab(MYCAURL);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("No input parameters...");
					}
				}
			}

			// Select the value from the dropdown by index
			else if (action.equalsIgnoreCase("select")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("") && intData != 0) {
						new Desktop_ActionKeywords().select(objectName, objectType, intData);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Invalid select...");
					}
				} 
			}

			// Select the value from the dropdown by value
			else if (action.equalsIgnoreCase("selectByValue")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")
							&& !data.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().selectByValue(objectName, objectType, data);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Invalid select...");
					}
				} 
			}

			else if (action.equalsIgnoreCase("selectByVisibleText")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")
							&& !data.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().selectByVisibleText(objectName, objectType, data);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Invalid select...");
					}
				} 
			}

			else if (action.equalsIgnoreCase("Highlight")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().Highlight_Object(objectName, objectType);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					}
				} 
				else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Invalid select...");
				}
			} 
			

			else if (action.equalsIgnoreCase("back")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().backButton();
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				}  else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
				}
			}

			else if (action.equalsIgnoreCase("close")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().closeBrowser();
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				}  else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
				}
			}

			// Quit Webdriver and all browser instances
			else if (action.equalsIgnoreCase("quit")) {
				if (OS.equalsIgnoreCase("Browser")) {
					new Desktop_ActionKeywords().quitBrowser();
					updateStatus(TCId, TSId, steprow, new Constants().pass);
					Thread.sleep(1000);
				}  else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
				}
			}

			else if (action.equalsIgnoreCase("wait")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (intData != 0) {
						// document.add(new Paragraph("wait"));
						new Desktop_ActionKeywords().wait(intData);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Time not provided...");
					}
				} 
			}

			else if (action.equalsIgnoreCase("scrollIntoView")) {
				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						new Desktop_ActionKeywords().scroll_view(objectName, objectType);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Not scrolled...");
					}
				}  else {
					System.out.println("Please cahnge the OS Type");
				}
			}

			else if (action.equalsIgnoreCase("hover")) {
				if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
					new Desktop_ActionKeywords().mouseHover(objectName, objectType);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("No input parameters...");
				}
			}

			// Perform cookie clear action
			else if (action.equalsIgnoreCase("clearcookies")) {
				new Desktop_ActionKeywords().clearCookies();
				updateStatus(TCId, TSId, steprow, new Constants().pass);
			}

			// Expands transactions in estmt page
			else if (action.equalsIgnoreCase("expand")) {
				if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
					new Desktop_ActionKeywords().expand(objectName, objectType);
					updateStatus(TCId, TSId, steprow, new Constants().pass);
				} else {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Invalid details...");
				}
			}

			// Action to navigate to current frame
			else if (action.equalsIgnoreCase("frame")) {

				if (OS.equalsIgnoreCase("Browser")) {
					if (!objectName.equalsIgnoreCase("") && !objectType.equalsIgnoreCase("")) {
						// new Desktop_ActionKeywords().selectFrame(objectName, objectType);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Invalid details...");
					}
				} 
			}

			// Action to navigate to frame using index
			else if (action.equalsIgnoreCase("selectFrameIndex")) {

				if (OS.equalsIgnoreCase("Browser")) {
					if (intData != 0) {
						new Desktop_ActionKeywords().selectFrameIndex(intData);
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} else {
						updateStatus(TCId, TSId, steprow, new Constants().fail);
						System.out.println("Invalid details...");
					}
				}
			}

			// Action to navigate to child window
			else if (action.equalsIgnoreCase("windowChild")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().selectWindowChild();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}
			

			else if (action.equalsIgnoreCase("refresh")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().refresh();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					}
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}

			else if (action.equalsIgnoreCase("windowParent")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().ParentWindow();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}

			// Action to navigate to parent window
			else if (action.equalsIgnoreCase("SwitchwindowParent")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().selectWindowParent();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to parent window");
				}
			}

			// Action to navigate to child window
			// else if(action.equalsIgnoreCase("SwitchFrame"))
			else if (action.equalsIgnoreCase("SwitchChildWindow")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().selectWindowChild();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}

			// Action to accept the popup
			else if (action.equalsIgnoreCase("acceptPopup")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().acceptPopup();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to accept the pop up");
				}
			}

			// Action to dismiss the popup
			else if (action.equalsIgnoreCase("dismissPopup")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().dismissPopup();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to dismiss the pop up");
				}
			}

			else if (action.equalsIgnoreCase("dismissPopup")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().dismissPopup();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to dismiss the pop up");
				}
			}

			
			else if (action.equalsIgnoreCase("scrollDownslow")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().scrollDown();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}

			else if (action.equalsIgnoreCase("scrollTop")) {
				try {
					if (OS.equalsIgnoreCase("Browser")) {
						new Desktop_ActionKeywords().scrollTop();
						updateStatus(TCId, TSId, steprow, new Constants().pass);
					} 
				} catch (Exception e) {
					updateStatus(TCId, TSId, steprow, new Constants().fail);
					System.out.println("Unable to switch to child window");
				}
			}


			else if (action.equalsIgnoreCase("Enter")) {
				new Desktop_ActionKeywords().PressEnterButton();
				updateStatus(TCId, TSId, steprow, new Constants().pass);
			}

			else if (action.equalsIgnoreCase("Tab")) {
				new Desktop_ActionKeywords().PressTABButton();
				updateStatus(TCId, TSId, steprow, new Constants().pass);
			}

			else if (action.equalsIgnoreCase(" ")) {
				updateStatus(TCId, TSId, steprow, new Constants().fail);
				System.out.println("Invalid Action");
			} else {
				updateStatus(TCId, TSId, steprow, new Constants().fail);
				System.out.println("Invalid Action");
			}
		}

		catch (Exception e) {
			updateStatus(TCId, TSId, steprow, new Constants().fail);
		}
	}

	public void updateStatus(String TCID, String TSID, XSSFRow statusupdate, String result) throws Exception {
		fis.close();
		System.out.println(result);
		addIntoList(TCID, TSID, result);
		MainScript.Maindocument.add(new Paragraph(TCID + "-" + TSID + "-" + result + " -- " + description));
		workbook = new XSSFWorkbook(new FileInputStream(new Constants().testCaseFile));
	}

	public void addIntoList(String TCID, String TSID, String result) {
		String name = TCID + "." + TSID;
		StepWiseResult.put(name, result);

		if (name.contains(TCID) && result.equalsIgnoreCase("Pass")) {
			TestCasePass.put(TCID, result);
		} else if (name.contains(TCID) && result.equalsIgnoreCase("Fail")) {
			TestCaseFail.put(TCID, result);
		}
	}

	public static void createTestReport(String time) throws Exception, DocumentException {
		ExcelUtils.TestCasePass.keySet().removeAll(ExcelUtils.TestCaseFail.keySet());
		Document document = new Document();
		PdfWriter.getInstance(document,
				new FileOutputStream(new File(new Constants().screenPath + "FinalTestResults.pdf")));
		document.open();
		document.add(new Paragraph("KDD AUTOMATION TEST RESULT REPORT"));
		document.add(new Paragraph("KDD Framework"));
		document.add(new Paragraph("Total Execution Time:   " + time));
		PdfPTable table = createTable1();
		table.setSpacingBefore(10);
		table.setSpacingAfter(5);
		document.add(table);
		if (ExcelUtils.TestCasePass.size() > 0) {
			table = createTable2();
			table.setSpacingBefore(10);
			table.setSpacingAfter(5);
			document.add(table);
		}
		if (ExcelUtils.TestCaseFail.size() > 0) {
			table = createTable3();
			table.setSpacingBefore(10);
			table.setSpacingAfter(5);
			document.add(table);
		}
		document.close();
	}

	public static PdfPTable createTable1() throws DocumentException {
		PdfPTable table = new PdfPTable(3);
		table.setWidthPercentage(500 / 6f);
		table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
		table.addCell("Total TestCase Executed");
		table.addCell("Passed");
		table.addCell("Failed");
		table.setHeaderRows(1);
		PdfPCell[] cells = table.getRow(0).getCells();
		for (int j = 0; j < cells.length; j++) {
			cells[j].setBackgroundColor(BaseColor.GRAY);
		}
		int total = ExcelUtils.TestCasePass.size() + ExcelUtils.TestCaseFail.size();
		String TotalNumbers = String.valueOf(total);
		String Passed = String.valueOf(ExcelUtils.TestCasePass.size());
		String Failed = String.valueOf(ExcelUtils.TestCaseFail.size());
		table.addCell(TotalNumbers);
		table.addCell(Passed);
		table.addCell(Failed);
		return table;
	}

	public static PdfPTable createTable2() throws DocumentException {
		PdfPTable table = new PdfPTable(2);
		table.setWidthPercentage(500 / 6f);
		table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
		table.addCell("Passed Testcase Number");
		table.addCell("Status");
		table.setHeaderRows(1);
		PdfPCell[] cells = table.getRow(0).getCells();
		for (int j = 0; j < cells.length; j++) {
			cells[j].setBackgroundColor(BaseColor.GREEN);
		}
		for (String s : ExcelUtils.TestCasePass.keySet()) {
			String value = ExcelUtils.TestCasePass.get(s);
			table.addCell(s);
			table.addCell(value);
		}
		return table;
	}

	public static PdfPTable createTable3() throws DocumentException {
		PdfPTable table = new PdfPTable(2);
		table.setWidthPercentage(500 / 6f);
		table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
		table.addCell("Failed Testcase Number");
		table.addCell("Status");
		table.setHeaderRows(1);
		PdfPCell[] cells = table.getRow(0).getCells();
		for (int j = 0; j < cells.length; j++) {
			cells[j].setBackgroundColor(BaseColor.RED);
		}
		for (String s : ExcelUtils.TestCaseFail.keySet()) {
			String value = ExcelUtils.TestCaseFail.get(s);
			table.addCell(s);
			table.addCell(value);
		}
		return table;
	}
}
