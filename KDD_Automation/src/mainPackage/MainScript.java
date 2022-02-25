package mainPackage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

public class MainScript {
	Date startTime, endTime;
	long timeElapsed;
	String totalTimeElapsed;
	public static Document Maindocument;
	public static OutputStream Mainfile;

	public static void main(String[] args) throws Exception {
		new MainScript();
	}

	// This is the main script, in which below actions are getting initiated.
	// 1. TestResults.pdf file is created for test execution logs will be generated
	// for each test steps
	// 2. Once all the test steps are executed then it should create
	// FinalTestResults.pdf with testcase details
	public MainScript() throws Exception {
		try {
			Mainfile = new FileOutputStream(new File(new Constants().screenPath + "Testresults.pdf"));
			Maindocument = new Document();
			PdfWriter.getInstance(Maindocument, Mainfile);
			Maindocument.open();
			startTime = new Date();
			new ExcelUtils().readTestcase();
			endTime = new Date();
			timeElapsed = endTime.getTime() - startTime.getTime();
			totalTimeElapsed = (timeElapsed / (60 * 60 * 1000) % 24) + "" + (timeElapsed / (60 * 1000) % 60) + " Min:"
					+ (timeElapsed / 1000 % 60 + " Sec");
			Maindocument.add(new Paragraph("Time Elapsed = " + (timeElapsed / (60 * 60 * 1000) % 24) + " "
					+ (timeElapsed / (60 * 1000) % 60) + " " + (timeElapsed / 1000 % 60) + " "));
			ExcelUtils.createTestReport(totalTimeElapsed);
			Maindocument.close();
			Mainfile.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

}
