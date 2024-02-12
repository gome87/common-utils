package com.lifeone.utils.poi2;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.GraphicsEnvironment;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PptxToPng {

	private static final Logger LOG = LoggerFactory.getLogger(PptxToPng.class);

	public void pptxToPng(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {

		FileInputStream fis = null;

		try {

			LOG.info("##### PPTX STEP 01 #####");

			GraphicsEnvironment e = GraphicsEnvironment.getLocalGraphicsEnvironment();
			String[] fontNames = e.getAvailableFontFamilyNames();

			LOG.info("##### PPTX STEP 02 #####");

			// pptx파일 찾기
			fis = new FileInputStream(new File(sOrgPath));
			XMLSlideShow ppt = new XMLSlideShow(fis);

			LOG.info("##### PPTX STEP 03 #####");

			// ppt 정보 찾기
			Dimension  sheet  = ppt.getPageSize();
			int        width  = sheet.width;
			int        height = sheet.height;
			int        count  = ppt.getSlides().size();

			LOG.info("##### PPTX STEP 04 #####");

			// 이미지 생성 정보 세팅
			BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = img.createGraphics();

			Font f = new Font("malgun", Font.ITALIC, 30);
			graphics.setFont(f);

			LOG.info("##### PPTX STEP 05 #####");

			int i = 1;
			for(XSLFSlide shape : ppt.getSlides()) {

				LOG.info("##### PPTX STEP 05-01 #####");

				// 화면 그리기
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, width, height));
				shape.draw(graphics);

				LOG.info("##### PPTX STEP 05-02 #####");

				// 파일 생성
				String sConvertFileName = sImgRootPath + "Test" + "_" + i + ".png";
				FileOutputStream fos = new FileOutputStream(new File(sConvertFileName));
				ImageIO.write(img, "PNG", fos);
				fos.close();
				i++;

				LOG.info("##### PPTX STEP 05-03 #####");

				// 이미지 경로 저장

			}

			LOG.info("##### PPTX STEP 06 #####");
		} catch (Exception e) {
			fis.close();
			LOG.error(">>>>>>>>>> pptxToPng = {}", e.getMessage());
			throw new Exception("##### pptxToPng Error 발생 #####");
		} finally {
			if(fis != null) {
				try {
					fis.close();
				} catch (Exception e2) {
					LOG.error(">>>>>>>>>> pptxToPng file close = {}", e2.getMessage());
					throw new Exception("##### pptxToPng File Close Error 발생 #####");
				}
			}
		}
	}

}
