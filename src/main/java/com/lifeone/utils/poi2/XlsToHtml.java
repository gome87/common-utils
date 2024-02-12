package com.lifeone.utils.poi2;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

public class XlsToHtml {

	private static final Logger LOG = LoggerFactory.getLogger(XlsToHtml.class);


	// Office(Excel) -> Html
	public void xlsToHtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {

		ByteArrayOutputStream outStream = null;

		try {
			LOG.info("##### XLS STEP 01 #####");

			// 대상 파일 존재 여부 확인
			InputStream input = new FileInputStream(sOrgPath);

			// 1. 엑셀 객체 생성
			HSSFWorkbook excelBook = new HSSFWorkbook(input);


			LOG.info("##### XLS STEP 02 #####");

			// 2. XLS를 HTML로 변경 설정
			ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
			excelToHtmlConverter.setOutputColumnHeaders(false);
			excelToHtmlConverter.processWorkbook(excelBook);
			Document htmlDocument = excelToHtmlConverter.getDocument();

			LOG.info("##### XLS STEP 03 #####");

			// 3. XLS를 HTML로 변환
			outStream = new ByteArrayOutputStream();
			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(outStream);
			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();

			LOG.info("##### XLS STEP 04 #####");

			serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			outStream.close();

			LOG.info("##### XLS STEP 05 #####");

			// 4. HTML 저장
			String content = new String(outStream.toByteArray());
			FileUtils.writeStringToFile(new File(sImgPath), content, "UTF-8");

			LOG.info("##### XLS STEP 06 #####");

		} catch (Exception e) {
			outStream.close();
			LOG.error(">>>>>>>>>> xlsToHtml = {}", e.getMessage());
			throw new Exception("##### xlsToHtml Error 발생 #####");
		}
	}

}
