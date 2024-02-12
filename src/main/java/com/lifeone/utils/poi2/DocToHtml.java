package com.lifeone.utils.poi2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

public class DocToHtml {

	private static final Logger LOG = LoggerFactory.getLogger(DocToHtml.class);

	// Office(DOC) -> Html
	public void word2003ToHtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception{

		OutputStream outputStream = null;

		try {
			// 대상 파일 존재 여부 확인
			File file = new File(sOrgPath);

			LOG.info("##### TEST STEP 01 #####");

			// 1. 워드파일을 읽어와서 HWPFDocument 객체 생성
			InputStream inputStream = new FileInputStream(file);

			LOG.info("##### TEST STEP 01-01 #####");

			HWPFDocument wordDocument = new HWPFDocument(inputStream);

			LOG.info("##### TEST STEP 01-02 #####");

			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());

			LOG.info("##### TEST STEP 02 #####");

			// 2. 이미지 저장 위치 설정
			wordToHtmlConverter.setPicturesManager(new PicturesManager() {
				public String savePicture(byte[] content, PictureType pictureType, String suggestedName,
						float widthInches, float heightInches) {
					File imgPath = new File(sImgPath);
					if (!imgPath.exists()) {
						imgPath.mkdirs();
					}
					File file = new File(sImgRootPath + suggestedName);
					try {
						OutputStream os = new FileOutputStream(file);
						os.write(content);
						os.close();
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
					return sChgPath + "/" + suggestedName;
				}
			});

			LOG.info("##### TEST STEP 03 #####");

			// 3. Word 문서 분석
			wordToHtmlConverter.processDocument(wordDocument);
			Document htmlDocument = wordToHtmlConverter.getDocument();

			outputStream = new FileOutputStream(sChgPath);

			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(outputStream);

			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer serializer = factory.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");

			serializer.transform(domSource, streamResult);


			LOG.info("##### TEST STEP 04 #####");

			// 4. File 작성 완료
			outputStream.close();

		} catch (Exception e) {
			//outputStream.close();
			LOG.error(">>>>>>>>>> word2003ToHtml = {}", e.getMessage());
			throw new Exception("##### word2003ToHtml Error 발생 #####");
		}
	}

}
