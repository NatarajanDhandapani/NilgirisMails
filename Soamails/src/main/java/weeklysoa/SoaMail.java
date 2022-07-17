package weeklysoa;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SoaMail {

	public static void main(String[] args) throws IOException, Exception, Exception {
		Map<Integer, String> add1 = new HashMap<Integer, String>();
		Map<Integer, String> add2 = new HashMap<Integer, String>();
		FileInputStream xlsfile = new FileInputStream("e:/debtors/Master.xlsx");
		Workbook xls = WorkbookFactory.create(xlsfile);
		Sheet sht = xls.getSheetAt(2);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				add1.put((int) r.getCell(0).getNumericCellValue(), r.getCell(2).getStringCellValue());
				add2.put((int) r.getCell(0).getNumericCellValue(), r.getCell(3).getStringCellValue());
			}
		}
		xlsfile.close();
		int code = 0;
		File[] soa = new File("e://mails/").listFiles();
		for (File s : soa) {
			PDDocument soamail = Loader.loadPDF(s);
			{
				PDFTextStripper soatext = new PDFTextStripper();
				soatext.setStartPage(1);
				soatext.setEndPage(1);
				String pdfFileInText = soatext.getText(soamail);
				String lines[] = pdfFileInText.split("\\r?\\n");
				code = Integer.parseInt(lines[4].substring(lines[4].indexOf(":") + 2));

				System.out.println(lines[4].substring(lines[4].indexOf(":") + 2) + " ~ "
						+ lines[5].substring(lines[5].indexOf(":") + 2) + " ~ "
						+ lines[16].substring(lines[16].indexOf(":") + 2, 43) + " ~ " + add1.get(code) + " ~ "
						+ add2.get(code));

			}
		}

		System.out.println("Dhandapani ... completed ");
		 
	}

}
