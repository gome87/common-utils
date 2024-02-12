package com.lifeone.utils.poi2;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;

public class DocxToPdf {

	private static final Logger LOG = LoggerFactory.getLogger(DocxToPdf.class);

	// https://github.com/eugenp/tutorials/blob/master/text-processing-libraries-modules/pdf/src/main/java/com/baeldung/pdf/DocxToPDFExample.java
	public void word2007ToPdf(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {

		OutputStream pdfOutputStream = null;
		Document pdfDocument = null;

		try {

			InputStream docxInputStream = new FileInputStream(sOrgPath);
			XWPFDocument document = new XWPFDocument(docxInputStream);
			pdfOutputStream = new FileOutputStream(sChgPath);

			pdfDocument = new Document(PageSize.A4, 50, 50, 50, 50);
			PdfWriter.getInstance(pdfDocument, pdfOutputStream);
			pdfDocument.open();

			BaseFont objBaseFont = BaseFont.createFont("font/malgun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
			Font objFont = new Font(objBaseFont);

			List<XWPFParagraph> paragraphs = document.getParagraphs();
			for (XWPFParagraph paragraph : paragraphs) {
				pdfDocument.add(new Paragraph(paragraph.getText(), objFont));
			}
			pdfDocument.close();
		} catch (Exception e) {
			pdfOutputStream.close();
			pdfDocument.close();
			LOG.error(">>>>>>>>>> word2007ToPdf = {}", e.getMessage());
			throw new Exception("##### word2007ToPdf Error 발생 #####");
		}
	}

}
