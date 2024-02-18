package com.lifeone.utils.poi2;

import java.io.File;
import java.io.PrintWriter;
import java.io.Writer;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.fit.pdfdom.PDFDomTree;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PdfToHtml {

	private static final Logger LOG = LoggerFactory.getLogger(PdfToHtml.class);

	// pdf to html
	public void pdfToHtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception{
		try {
			PDDocument pdf = PDDocument.load(new File(sOrgPath));
		    Writer output = new PrintWriter(sChgPath, "utf-8");
		    new PDFDomTree().writeText(pdf, output);
		    output.close();

		    LOG.info("##### ##### Html created successfully.");
		} catch (Exception e) {
			LOG.error(">>>>>>>>>> pdfToHtml = {}", e.getMessage());
			throw new Exception("##### pdfToHtml Error 발생 #####");
		}
	}

}
