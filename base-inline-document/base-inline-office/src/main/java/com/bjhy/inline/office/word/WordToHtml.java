package com.bjhy.inline.office.word;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStreamWriter;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.License;
import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * word 转成html
 * @author wubo 
 */
public class WordToHtml {

	  
	//doc转换为html
	public void docToHtml(OfficeStorePath storePath) throws Exception { 
		//得到Html文件路径
		String htmlFilePath = storePath.getHtmlFilePath();
//		int i = 1/0;
		//得到图片目录
		final String imageDirectory = storePath.getImagesDirectory();
		//得到URIResolver目录
		final String uriResolverDirectory = storePath.getURIResolverDirectory();
//		org.apache.poi.POIXMLTypeLoader
	    
	    HWPFDocument wordDocument = new HWPFDocument(storePath.getFileStream()); 
	    Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument(); 
	    WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document); 
		// 保存图片，并返回图片的相对路径
		wordToHtmlConverter.setPicturesManager(new PicturesManager() {
			public String savePicture(byte[] content, PictureType pictureType,String name, float width, float height) {
				try (FileOutputStream out = new FileOutputStream(imageDirectory +"/"+ name)) {
					out.write(content);
				} catch (Exception e) {
					e.printStackTrace();
				}
				return uriResolverDirectory +"/"+ name;
			}
		});
	    
	    wordToHtmlConverter.processDocument(wordDocument); 
	    Document htmlDocument = wordToHtmlConverter.getDocument(); 
	    DOMSource domSource = new DOMSource(htmlDocument); 
	    StreamResult streamResult = new StreamResult(new File(htmlFilePath)); 
	  
	    TransformerFactory tf = TransformerFactory.newInstance(); 
	    Transformer serializer = tf.newTransformer(); 
	    serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8"); 
	    serializer.setOutputProperty(OutputKeys.INDENT, "yes"); 
	    serializer.setOutputProperty(OutputKeys.METHOD, "html"); 
	    serializer.transform(domSource, streamResult); 
	  }
	
	

	public void docxToHtml(OfficeStorePath storePath) throws Exception {
		
		//得到Html文件路径
		String htmlFilePath = storePath.getHtmlFilePath();
		//得到图片目录
		String imageDirectory = storePath.getImagesDirectory();
		//得到URIResolver目录
		String uriResolverDirectory = storePath.getURIResolverDirectory();
		
		OutputStreamWriter outputStreamWriter = null;
		try {
			XWPFDocument document = new XWPFDocument(storePath.getFileStream());
			XHTMLOptions options = XHTMLOptions.create();
			// 存放图片的文件夹
			options.setExtractor(new FileImageExtractor(new File(imageDirectory)));
			// html中图片的路径
			options.URIResolver(new BasicURIResolver(uriResolverDirectory));
			outputStreamWriter = new OutputStreamWriter(new FileOutputStream(htmlFilePath), "utf-8");
			XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
			xhtmlConverter.convert(document, outputStreamWriter, options);
		} finally {
			if (outputStreamWriter != null) {
				outputStreamWriter.close();
			}
		}
	}
	
	/**
	 * 利用  aspose转html
	 * @param storePath
	 */
	public void asposeWordToHtml(OfficeStorePath storePath){
		//得到Html文件路径
		String htmlFilePath = storePath.getHtmlFilePath();
//		//得到图片目录
//		String imageDirectory = storePath.getImagesDirectory();
//		//得到URIResolver目录
//		String uriResolverDirectory = storePath.getURIResolverDirectory();
		
		loadLicense();//加载licence
		
		try {
			com.aspose.words.Document doc = new com.aspose.words.Document(storePath.getFileStream());
			HtmlSaveOptions saveOptions = new HtmlSaveOptions();
			saveOptions.setExportImagesAsBase64(true);
			doc.save(htmlFilePath, saveOptions);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 加载licence
	 * @return
	 */
	private boolean loadLicense(){
		InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("config/license.xml");
		boolean result = false;
		try {
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
}
