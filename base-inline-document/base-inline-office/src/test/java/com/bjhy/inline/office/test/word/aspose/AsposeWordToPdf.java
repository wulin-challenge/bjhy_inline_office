package com.bjhy.inline.office.test.word.aspose;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import com.aspose.words.SaveOptions;
import com.itextpdf.text.log.SysoCounter;

public class AsposeWordToPdf {
	private static String rootPath = System.getProperty("user.dir")+"/";
	private boolean getLicense() {
		boolean result = false;
		try {
			InputStream is = new FileInputStream(rootPath+"config"+File.separator+"license.xml");
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	private void wordToPdf() throws Exception{
		String sourceAddress = "F:/resources/temp/temp1/aspose/wordToPdf/2.docx";
		String targetAddress = "F:/resources/temp/temp1/aspose/wordToPdf/2.pdf";
		String targetAddressHtml = "F:/resources/temp/temp1/aspose/wordToPdf/2.html";
		getLicense();
		com.aspose.words.Document doc = new com.aspose.words.Document(sourceAddress); // Address是将要被转化的word文档
//		doc.setResourceLoadingCallback(value);
		doc.save(targetAddress, SaveFormat.PDF);//
		HtmlSaveOptions saveOptions = new HtmlSaveOptions();
		saveOptions.setExportImagesAsBase64(true);

		doc.save(targetAddressHtml, saveOptions);
//		doc.save(targetAddressHtml, SaveFormat.HTML);//
		
		
	}

	public static void main(String[] args) throws Exception {
		AsposeWordToPdf asposeWordToPdf = new AsposeWordToPdf();
		asposeWordToPdf.wordToPdf();
	}
}
