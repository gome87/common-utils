package com.lifeone.utils.poi2;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.tidy.Tidy;
import org.xhtmlrenderer.pdf.ITextRenderer;

import com.itextpdf.text.pdf.BaseFont;
import com.lowagie.text.DocumentException;


public class XhtmlToPdf {

	private static final Logger LOG = LoggerFactory.getLogger(XhtmlToPdf.class);


	// Html -> Xhtml
	// https://blog.naver.com/dlrhkdgh3333/223354207241
	public void htmlToXhtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {

		// Tidy 객체 생성
		Tidy tidy = new Tidy();

		// 인코딩 설정
		tidy.setInputEncoding("UTF-8");
		tidy.setOutputEncoding("UTF-8");
		tidy.setDocType("omit");

		// XHTML 출력 설정
		tidy.setXHTML(true);

		// 변환 실행
		try {
			InputStream inputStream = new FileInputStream(sOrgPath);
			OutputStream outputStream = new FileOutputStream(sChgPath);
			tidy.parse(inputStream, outputStream);
			LOG.info("##### XHTML created successfully.");
		} catch (IOException e) {
			throw new RuntimeException("HTML을 XHTML로 변환하는 중 오류 발생", e);
		}
		LOG.info("##### HTML을 XHTML로 성공적으로 변환했습니다.");
	}

	// Xhtml -> PDF
	public void xhtmlToPdf(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {
		try {
			OutputStream os = new FileOutputStream(sChgPath);

			// PDF 생성을 위한 ITextRenderer 인스턴스 생성
			ITextRenderer renderer = new ITextRenderer();

			renderer.getFontResolver().addFont(("font/malgun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

			// XHTML 파일을 ITextRenderer에 설정
			renderer.setDocument(new File(sOrgPath));

			// PDF 생성
			renderer.layout();
			renderer.createPDF(os);
			renderer.finishPDF(); // PDF 생성 완료

			LOG.info("##### PDF created successfully.");
		} catch (DocumentException | IOException e) {
			e.printStackTrace();
		}
	}

}
