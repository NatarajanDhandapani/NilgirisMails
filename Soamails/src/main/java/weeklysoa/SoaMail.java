// updated on 18.04.23 @ 08.10 hrs

package weeklysoa;
import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
public class SoaMail {
	public static void main(String[] args) throws IOException, Exception, Exception {
		
		int code = 0;
		File[] soa = new File("d://soa data/mails/").listFiles();
		for (File s : soa) {
			PDDocument soamail = Loader.loadPDF(s);
			{
				PDFTextStripper soatext = new PDFTextStripper();
				soatext.setStartPage(1);
				soatext.setEndPage(1);
				String pdfFileInText = soatext.getText(soamail);
				String lines[] = pdfFileInText.split("\\r?\\n");
				for (int a = 0; a < 45; a++) {
					if (lines[a].contains("Ship To"))
						code = Integer.parseInt(lines[a].substring(0, 6));
				}
				File f = new File("d://soa data/mails/" + s.getName());
				f.renameTo(new File("d://soa data/mails/" + code + ".pdf"));
			}
			System.out.println(s);
		}
	}
}
