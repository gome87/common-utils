package com.lifeone.utils.poi;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.imageio.ImageIO;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xhtmlrenderer.simple.Graphics2DRenderer;
import org.xhtmlrenderer.simple.XHTMLPanel;

public class HancomOfficeConverter {

	private static final Logger LOG = LoggerFactory.getLogger(HancomOfficeConverter.class);

	// https://copyprogramming.com/howto/paginating-xhtml-to-png-image-files-using-flying-saucer#paginating-xhtml-to-png-image-files-using-flying-saucer

	/**
	 * converter XHTML을 Image 변환
	 *
	 * @param String sFilePath  파일경로
	 * @param String sImageDir  이미지 폴더 경로
	 * @return String 이미지 변환 경로
	 * @throws Exception
	 * @since 2024. 02. 28
	 * @author 김영우
	 */
	public String convertXhtmlToImage(String sFilePath, String sImageDir) throws Exception {

		InputStream is = null;

		try {
			is = new FileInputStream(sFilePath) ;

			File file = File.createTempFile(sImageDir, ".png");

			XHTMLPanel panel = new XHTMLPanel();
			panel.setSize(1024, 768);
			panel.setDocument(is, "");

			BufferedImage img = new BufferedImage(1024, 768, BufferedImage.TYPE_INT_ARGB);
			Graphics2D graphics = (Graphics2D) img.getGraphics();
			graphics.setColor(Color.white);
			graphics.fillRect(0, 0, 1024, 768);

			Graphics2DRenderer renderer = new Graphics2DRenderer();
			renderer.setDocument(panel.getDocument(), "");
			renderer.layout(graphics, null);
			renderer.render(graphics);

			ImageIO.write(img, "png", file);

		} catch (Exception e) {
			LOG.error(">>>>>>>>>> convertXhtmlToImage = {}", e.getMessage());
			throw new Exception("##### convertXhtmlToImage Error 발생 #####");
		} finally {
			if(is != null) {
				is.close();
			}
		}

		return null;
	}
}
