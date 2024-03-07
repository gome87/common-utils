package com.lifeone.utils.poi;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PowerPointConverter {

	private static final Logger LOG = LoggerFactory.getLogger(PowerPointConverter.class);

	// https://blog.csdn.net/qq_38567039/article/details/88418965
	// https://blog.csdn.net/m0_37615697/article/details/81083986

	/**
	 * converter ppt2003/2007(ppt/Pptx)을 HTML 변환
	 *
	 * @param String sFilePath      파일경로
	 * @param String sHtmlImageDir  HTML 이미지 경로
	 * @param String sHtmlPath      HTML 파일 경로
	 * @return String 이미지 변환 경로
	 * @throws Exception
	 * @since 2024. 03. 08
	 * @author 김영우
	 */
	public List<Map<String, Object>>  convertPptToHtml(String sFilePath, String sHtmlImageDir, String sHtmlPath) throws Exception {

		List<Map<String, Object>> addList = new ArrayList<>();
		Map<String, Object>       addMap  = null;
		String htmlStr = "";

		// 파일 읽기
		File pptFile = new File(sFilePath);

		if(pptFile.exists()){
			try {
				// 파일 확장자 찾기
				String sType = sFilePath.substring(sFilePath.lastIndexOf(".") + 1);
				if(StringUtils.isNotBlank(sType) && "ppt".equals(sType)) {
					htmlStr = this.toImage2003(sFilePath, sHtmlImageDir, sHtmlPath);
				} else if(StringUtils.isNotBlank(sType) && "pptx".equals(sType)) {
					htmlStr = this.toImage2007(sFilePath, sHtmlImageDir, sHtmlPath);
				} else {
					LOG.error("##### 변환 대상 확장자가 아닙니다. #####");
					return null;
				}
			} catch (Exception e) {
				LOG.error("##### PPT/PPTX 변환 중 에러가 발생하였습니다. #####");
			}
		} else {
			return null;
		}

		if(StringUtils.isNotBlank(htmlStr)) {
			// 파일 디렉토리 생성
			Files.createDirectories(Paths.get(sHtmlPath));

			// 파일 생성
			String sTargetFilePath = sHtmlPath.concat("index.html");
			File targetFile = new File(sTargetFilePath);
			if(!targetFile.exists()) {
				targetFile.createNewFile();
			}
			FileUtils.writeStringToFile(targetFile, htmlStr, "UTF-8");

			addMap = new HashMap<>();
			addMap.put("FILE_URL", sTargetFilePath);
			addList.add(addMap);
		}
		return addList;
	}

	// PPT -> IMAGE(PNG)
	private String toImage2003(String sSourcePath, String sHtmlImageDir, String sTaragetDir) {
		String       sHtmlStr = "";
		StringBuffer sb       = null;

		try {
			// 파일 읽기
			HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(sSourcePath));

			// 파일 디렉토리 생성
			Files.createDirectories(Paths.get(sTaragetDir));

			sb = new StringBuffer();

			Dimension pgsize = ppt.getPageSize();

			for(int i=0; i<ppt.getSlides().size(); i++) {
				for(HSLFShape shape : ppt.getSlides().get(i).getShapes()) {
					if(shape instanceof HSLFTextShape) {
						HSLFTextShape tsh = (HSLFTextShape) shape;
						for(HSLFTextParagraph p : tsh) {
							for(HSLFTextRun r : p) {
								r.setFontFamily("Noto Sans");
							}
						}
					}
				}

				BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();

				// 작성화면 초기화
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

				// render
				ppt.getSlides().get(i).draw(graphics);

				// 이미지 경로 만들기
				String sImageDir = sHtmlImageDir;
				Files.createDirectories(Paths.get(sImageDir));

				// 이미지 파일명
				String sRelativeImagePath = String.valueOf((i + 1)).concat(".png");

				// 이미지 전체경로
				String sImagePath = sImageDir.concat(sRelativeImagePath);

				// 파일 생성
				FileOutputStream out = new FileOutputStream(sImagePath);
				ImageIO.write(img, "png", out);
				out.close();

				// BASE64 변경

				// 이미지 경로 세팅
				sb.append("<br>");
				sb.append("<img src=\"" + sImagePath + "\""+ " />");
			}

		} catch (Exception e) {
			LOG.error("##### PPT 이미지 변환 중 에러가 발생하였습니다. #####");
		}
		sHtmlStr = sb.toString();

		return null;
	}

	// PPTX -> IMAGE(PNG)
	private String toImage2007(String sSourcePath, String sHtmlImageDir, String sTaragetDir) {

		String       sHtmlStr = "";
		StringBuffer sb       = null;

		try {
			FileInputStream is = new FileInputStream(sSourcePath);
			XMLSlideShow ppt = new XMLSlideShow(is);
			is.close();

			// 파일 디렉토리 생성
			Files.createDirectories(Paths.get(sTaragetDir));

			sb = new StringBuffer();

			Dimension pgsize = ppt.getPageSize();

			for(int i = 0; i < ppt.getSlides().size(); i++) {
				// PPTX 폰트 설정
				for(XSLFShape shape : ppt.getSlides().get(i).getShapes()) {
					if(shape instanceof XSLFTextShape) {
						XSLFTextShape tsh = (XSLFTextShape) shape;
						for(XSLFTextParagraph p : tsh) {
							for(XSLFTextRun r : p) {
								r.setFontFamily("Noto Sans");
							}
						}
					}
				}

				BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();

				// 작성화면 초기화
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

				// render
				ppt.getSlides().get(i).draw(graphics);

				// 이미지 경로 만들기
				String sImageDir = sHtmlImageDir;
				Files.createDirectories(Paths.get(sImageDir));

				// 이미지 파일명
				String sRelativeImagePath = String.valueOf((i + 1)).concat(".png");

				// 이미지 전체경로
				String sImagePath = sImageDir.concat(sRelativeImagePath);

				// 파일 생성
				FileOutputStream out = new FileOutputStream(sImagePath);
				ImageIO.write(img, "png", out);
				out.close();

				// BASE64 변경

				// 이미지 경로 세팅
				sb.append("<br>");
				sb.append("<img src=\"" + sImagePath + "\""+ " />");
			}
		} catch (Exception e) {
			LOG.error("##### PPTX 이미지 변환 중 에러가 발생하였습니다. #####");
		}
		sHtmlStr = sb.toString();

		return sHtmlStr;
	}

	// Image 사이즈 조정
	private void resizeImage(String sSrcImgPath, String sDistImgPath, int nWidth, int nHeight){

		BufferedImage buffImg = null;

		try {
			// 이미지 읽기
			File srcFile = new File(sSrcImgPath);
			Image srcImg = ImageIO.read(srcFile);

			buffImg = new BufferedImage(nWidth, nHeight, BufferedImage.TYPE_INT_RGB);
			buffImg.getGraphics().drawImage(srcImg.getScaledInstance(nWidth, nHeight, Image.SCALE_SMOOTH), 0, 0, null);
			ImageIO.write(buffImg, "JPEG", new File(sDistImgPath));
		} catch (Exception e) {
			LOG.error("##### 이미지 사이즈 변경 중 에러가 발생하였습니다. #####");
		}
	}

}
