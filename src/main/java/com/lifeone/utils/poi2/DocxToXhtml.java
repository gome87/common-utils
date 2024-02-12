package com.lifeone.utils.poi2;

import java.io.OutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DocxToXhtml {

	private static final Logger LOG = LoggerFactory.getLogger(DocxToXhtml.class);

	// Office(DOCX) -> Xhtml
	public void word2007ToXhtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {

		OutputStream outputStream = null;

//		try {
//
//			LOG.info("#### DOCX TEST 01 #####");
//
//			// 대상 파일 존재 여부 확인
//			File file = new File(sOrgPath);
//
//			LOG.info("#### DOCX TEST 02 #####");
//
//			// 1. 워드파일을 읽어와서 XWPFDocument 객체 생성
//			InputStream inputStream = new FileInputStream(file);
//			XWPFDocument document = new XWPFDocument(inputStream);
//
//			LOG.info("#### DOCX TEST 03 #####");
//
//			// 2. XHTML 설정 세팅
//			XHTMLOptions options = XHTMLOptions.create();
//			options.URIResolver(new BasicURIResolver(sImgRootPath));
//			FileImageExtractor extractor = new FileImageExtractor(new File(sImgPath));
//			options.setExtractor(extractor);
//
//			LOG.info("#### DOCX TEST 04 #####");
//
//			// 3. 워드를 XHTML 변환
//			File htmlFile = new File(sChgPath);
//			outputStream = new FileOutputStream(htmlFile);
//			XHTMLConverter.getInstance().convert(document, outputStream, options);
//
//			LOG.info("#### DOCX TEST 05 #####");
//
//			// 4. File 작성 종료
//			outputStream.close();
//
//			LOG.info("#### DOCX TEST 06 #####");
//
//		} catch (Exception e) {
//			//outputStream.close();
//			LOG.error(">>>>>>>>>> word2007ToXhtml = {}", e.getMessage());
//			throw new Exception("##### word2007ToXhtml Error 발생 #####");
//		} finally {
//			if(outputStream != null) {
//				outputStream.close();
//			}
//		}
	}

}
