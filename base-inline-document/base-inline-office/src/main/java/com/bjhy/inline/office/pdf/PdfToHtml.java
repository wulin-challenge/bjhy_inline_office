package com.bjhy.inline.office.pdf;

import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;

import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * 将pdf转换为html
 * @author wubo
 *
 */
public class PdfToHtml {
	
	/**
	 * 将pdf转换为图片
	 * @param storePath
	 * @return
	 */
	public String pdfToHtml(OfficeStorePath storePath){
		StringBuffer html = new StringBuffer();
		int pdfNumber = pdf2img(storePath);//PDFBOX转图片,返回 共转换了多少张ppt

		if(pdfNumber != 0){
			for (int i = 1; i <=pdfNumber; i++) {
				String imagePath = storePath.getURIResolverDirectory()+"/"+i+".png";
				html.append("<br/>&nbsp;<img src="+imagePath+"><br/><br/>");
			}
		}else{
			html.append("");
		}
		return html.toString();
	}
	
	/** 
     *  PDFBOX转图片,返回 共转换了多少张ppt
     * @param pdfPath pdf文件的路径 
     * @param savePath 图片保存的地址 
     * @param imgType 图片保存方式 
     */  
    @SuppressWarnings("unchecked")
    public int pdf2img(OfficeStorePath storePath){  
    
    	//pdf的页数
    	int pdfNumber = 0;
        PDDocument pdDocument = null;   
        try {  
            PDFParser parser = new PDFParser(storePath.getFileStream());  
            parser.parse();  
            pdDocument = parser.getPDDocument();  
			List<PDPage> pages = pdDocument.getDocumentCatalog().getAllPages();  
			pdfNumber = pages.size();
			
            for (int i = 0; i < pages.size(); i++) {  
                String saveFileName = storePath.getImagesDirectory()+"/" + (i+1) +".png";  
                PDPage page =  pages.get(i);  
                pdfPage2Img(page,saveFileName);  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }finally{  
            if(pdDocument != null){  
                try {  
                    pdDocument.close();  
                } catch (IOException e) {  
                    e.printStackTrace();  
                }  
            }  
        }  
        return pdfNumber;
    }  
      
    /** 
     * pdf页转换成图片 
     * @param page       
     * @param saveFileName 
     * @throws IOException 
     */  
    public void pdfPage2Img(PDPage page,String saveFileName) throws IOException{
    	ImageWriter writer = null;
    	ImageOutputStream imageout = null;
        try {
			BufferedImage img_temp  = page.convertToImage();  
			Iterator<ImageWriter> it = ImageIO.getImageWritersBySuffix("png");  
			writer = (ImageWriter) it.next();   
			imageout = ImageIO.createImageOutputStream(new FileOutputStream(saveFileName));  
			writer.setOutput(imageout);  
			writer.write(new IIOImage(img_temp, null, null));
		} catch (Exception e) {
			e.printStackTrace();
		} finally{
			imageout.close();
		}
    }  
}
