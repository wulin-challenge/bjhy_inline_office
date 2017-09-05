package com.bjhy.inline.office.image;

import java.awt.image.BufferedImage;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.imageio.ImageIO;

import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * 将图片转成html
 * @author wubo
 */
public class ImageToHtml {
	
	public String imageToHtml(OfficeStorePath storePath) {
		InputStream fileStream = storePath.getFileStream();
		String fileName = storePath.getFileName();
		String suffix = storePath.getSuffix();
		
		StringBuffer html = new StringBuffer();

		String filename2 = storePath.getImagesDirectory() + "/" + fileName+ "." + suffix;

		try {
			BufferedImage image = ImageIO.read(fileStream);
			FileOutputStream out = new FileOutputStream(filename2);
			javax.imageio.ImageIO.write(image, suffix, out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		String imageUrl = storePath.getURIResolverDirectory()+ "/" + fileName+ "." + suffix;
		html.append("<br/>&nbsp;<img src="+imageUrl+"><br/><br/>");
		return html.toString();
	}

}
