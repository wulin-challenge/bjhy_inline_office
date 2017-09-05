package com.bjhy.inline.office.ppt;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * 这里其实是ppt转换图片 然后可以直接播放图片 动态效果没有了
 * @author wubo
 */
public class Ppt2Html {

	/**
	 * 将ppt转换为一张张图片
	 * @param storePath
	 * @return 共转换了多少张ppt
	 */
	private int doPPTtoImage(OfficeStorePath storePath) {
		int pptNumber = 0;
		String suffix = storePath.getSuffix();
		if("ppt".equalsIgnoreCase(suffix)){
			pptNumber = toImage2003(storePath);
		}else if("pptx".equalsIgnoreCase(suffix)){
			try {
				pptNumber = toImage2007(storePath);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
		return pptNumber;
	}
	
	private int toImage2003(OfficeStorePath storePath){  
		int pptNumber = 0;
        try {  
            HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(storePath.getFileStream()));  
              
            Dimension pgsize = ppt.getPageSize(); 
            pptNumber =  ppt.getSlides().size(); //得到ppt张数
            for (int i = 0; i < ppt.getSlides().size(); i++) {  
                //防止中文乱码  
                for(HSLFShape shape : ppt.getSlides().get(i).getShapes()){  
                    if(shape instanceof HSLFTextShape) {  
                        HSLFTextShape tsh = (HSLFTextShape)shape;  
                        for(HSLFTextParagraph p : tsh){  
                            for(HSLFTextRun r : p){  
                                r.setFontFamily("宋体");  
                            }  
                        }  
                    }  
                }  
                BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);  
                Graphics2D graphics = img.createGraphics();  
                // clear the drawing area  
                graphics.setPaint(Color.white);  
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));  
                  
                // render  
                ppt.getSlides().get(i).draw(graphics);  
                  
                // save the output  
                String filename = storePath.getImagesDirectory()+"/" + (i+1) + ".png";  
                System.out.println(filename);  
                FileOutputStream out = new FileOutputStream(filename);  
                javax.imageio.ImageIO.write(img, "png", out);  
                out.close();  
//              resizeImage(filename, filename, width, height);  
                  
            }  
            System.out.println("3success");  
        } catch (Exception e) {  
        	e.printStackTrace();
        }  
        return pptNumber;
    }  
	
	private int toImage2007(OfficeStorePath storePath) throws Exception{
		int pptNumber = 0;
        InputStream is = storePath.getFileStream();  
        XMLSlideShow ppt = new XMLSlideShow(is);  
        is.close();  
          
        Dimension pgsize = ppt.getPageSize();  
        System.out.println(pgsize.width+"--"+pgsize.height);  
          
        pptNumber =  ppt.getSlides().size(); //得到ppt张数
        for (int i = 0; i < ppt.getSlides().size(); i++) {  
            try {  
                //防止中文乱码  
                for(XSLFShape shape : ppt.getSlides().get(i).getShapes()){  
                    if(shape instanceof XSLFTextShape) {  
                        XSLFTextShape tsh = (XSLFTextShape)shape;  
                        for(XSLFTextParagraph p : tsh){  
                            for(XSLFTextRun r : p){  
                                r.setFontFamily("宋体");  
                            }  
                        }  
                    }  
                }  
                BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);  
                Graphics2D graphics = img.createGraphics();  
                // clear the drawing area  
                graphics.setPaint(Color.white);  
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));  
                  
                // render  
                ppt.getSlides().get(i).draw(graphics);  
                  
                // save the output  
                String filename = storePath.getImagesDirectory()+"/" + (i+1) + ".png";  
                System.out.println(filename);  
                FileOutputStream out = new FileOutputStream(filename);  
                javax.imageio.ImageIO.write(img, "png", out);  
                out.close();  
            } catch (Exception e) {  
                System.out.println("第"+i+"张ppt转换出错");  
            }  
        }  
        System.out.println("7success");  
        return pptNumber;
	}  
	
	public String pptToHtml(OfficeStorePath storePath){
		StringBuffer html = new StringBuffer();
		int pptNumber = doPPTtoImage(storePath);//将ppt转换为一张张图片,返回 共转换了多少张ppt

		if(pptNumber != 0){
			
			for (int i = 1; i <=pptNumber; i++) {
				String imagePath = storePath.getURIResolverDirectory()+"/"+i+".png";
				html.append("<br/>&nbsp;<font size='6'>第"+i+"页</font>&nbsp;<br/><img src="+imagePath+"><br/><br/>");
			}
			
		}else{
			html.append("");
		}
		return html.toString();
	}
	
	

}