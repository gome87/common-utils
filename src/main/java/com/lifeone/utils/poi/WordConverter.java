package com.lifeone.utils.poi;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import fr.opensagres.poi.xwpf.converter.core.BasicURIResolver;
import fr.opensagres.poi.xwpf.converter.core.FileImageExtractor;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;

public class WordConverter {

	private static final Logger LOG = LoggerFactory.getLogger(WordConverter.class);

	// https://blog.csdn.net/qq_38567039/article/details/88418965
	// https://blog.csdn.net/m0_37615697/article/details/81084018?ops_request_misc=%257B%2522request%255Fid%2522%253A%2522170982775616800180617186%2522%252C%2522scm%2522%253A%252220140713.130102334.pc%255Fall.%2522%257D&request_id=170982775616800180617186&biz_id=0&utm_medium=distribute.pc_search_result.none-task-blog-2~all~first_rank_ecpm_v1~rank_v31_ecpm-4-81084018-null-null.142^v99^control&utm_term=com.maiyue.base&spm=1018.2226.3001.4187

	/**
	 * converter word2003(doc)을 HTML 변환
	 *
	 * @param String sFilePath      파일경로
	 * @param String sHtmlImageDir  HTML 이미지 경로
	 * @param String sHtmlPath      HTML 파일 경로
	 * @return String 이미지 변환 경로
	 * @throws Exception
	 * @since 2024. 02. 23
	 * @author 김영우
	 */
	public String convertDocToHtml(String sFilePath, String sHtmlImageDir, String sHtmlPath) throws Exception {

		ByteArrayOutputStream outStream = null;

		try {
			// 워드 파일 읽기
			InputStream input = new FileInputStream(sFilePath);
			HWPFDocument doc = new HWPFDocument(input);

			// 파일 변환
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());

			// 변환 대상 이미지 이름 설정
			wordToHtmlConverter.setPicturesManager(new PicturesManager() {
				public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
						float heightInches) {
					return suggestedName;
				}
			});

			wordToHtmlConverter.processDocument(doc);

			// 변환 대상 이미지 저장 위치 설정
			List<Picture> pics = doc.getPicturesTable().getAllPictures();
			if(CollectionUtils.isNotEmpty(pics)) {
				for(Picture pic : pics) {
					pic.writeImageContent(new FileOutputStream(sHtmlImageDir.concat(pic.suggestFullFileName())));
				}
			}

			// HTML 변환
			Document htmlDocument = wordToHtmlConverter.getDocument();
			DOMSource domSource = new DOMSource(htmlDocument);

			outStream = new ByteArrayOutputStream();
			StreamResult streamResult = new StreamResult(outStream);

			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");

			serializer.transform(domSource, streamResult);
			outStream.close();

			String sContent = new String(outStream.toByteArray());

			// HTML 파일 생성
			FileUtils.writeStringToFile(new File(sHtmlPath), sContent, "utf-8");

		} catch (Exception e) {
			LOG.error(">>>>>>>>>> convertDocToHtml = {}", e.getMessage());
			throw new Exception("##### convertDocToHtml Error 발생 #####");
		} finally {
			if(outStream != null) {
				outStream.close();
			}
		}
		return sHtmlPath;
	}

	/**
	 * converter word2007(docx)을 HTML 변환
	 *
	 * @param String sFilePath      파일경로
	 * @param String sHtmlImageDir  HTML 이미지 경로
	 * @param String sHtmlPath      HTML 파일 경로
	 * @return String 이미지 변환 경로
	 * @throws Exception
	 * @since 2024. 02. 23
	 * @author 김영우
	 */
	public String convertDocxToHtml(String sFilePath, String sHtmlImageDir, String sHtmlPath) throws Exception {

		OutputStreamWriter osw = null;

		try {
			// 워드 파일 읽기
			XWPFDocument doc = new XWPFDocument(new FileInputStream(sFilePath));
			XHTMLOptions options = XHTMLOptions.create();

			// 이미지 경로
			options.setExtractor(new FileImageExtractor(new File(sHtmlImageDir)));
			options.URIResolver(new BasicURIResolver("image"));

			// 저장경로
			osw = new OutputStreamWriter(new FileOutputStream(sHtmlPath), "utf-8");

			// 변환
			XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
			xhtmlConverter.convert(doc, osw, options);

			// 파일 닫기
			osw.close();

		} catch (Exception e) {
			LOG.error(">>>>>>>>>> convertDocxToHtml = {}", e.getMessage());
			throw new Exception("##### convertDocxToHtml Error 발생 #####");
		} finally {
			if(osw != null) {
				osw.close();
			}
		}

		return sHtmlPath;
	}


}
