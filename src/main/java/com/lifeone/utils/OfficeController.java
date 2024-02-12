package com.lifeone.utils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.lifeone.utils.poi2.DocxToPdf;
import com.lifeone.utils.poi2.PptToPng;
import com.lifeone.utils.poi2.PptxToPng;
import com.lifeone.utils.poi2.XlsToHtml;
import com.lifeone.utils.poi2.XlsxToHtml;
import com.lifeone.utils.poi2.XlsxToPDF;

@Controller
public class OfficeController {

	private static final Logger LOG = LoggerFactory.getLogger(OfficeController.class);

	/**
	 * Office -> HTML/XHTML/PDF -> Image(PNG)
	 *
	 * @param HttpServletRequest  req
	 * @param HttpServletResponse rep
	 * @return
	 * @throws Exception
	 * @since 2024. 02. 12
	 * @author 김영우
	 */
	@RequestMapping(value = "/officeConverter", method = { RequestMethod.GET })
	public void retrieveOfficeConverter(HttpServletRequest req, HttpServletResponse res) throws Exception {

		LOG.info("##### Converter Test #####");

		//DocToHtml convert = new DocToHtml();
		//convert.word2003ToHtml("C:\\fileupload\\convert_after\\신청서(Application form).doc", "C:\\fileupload\\convert_image", "C:\\fileupload\\convert_image\\신청서(Application form).png", "C:\\fileupload\\convert_before\\신청서(Application form).html");

		//DocxToXhtml convert = new DocxToXhtml();
		//convert.word2007ToXhtml("C:\\fileupload\\convert_after\\안내문.docx", "C:\\fileupload\\convert_image", "C:\\fileupload\\convert_image\\안내문.png", "C:\\fileupload\\convert_before\\안내문.xhtml");

		// 가능성 있음
		//DocxToPdf convert = new DocxToPdf();
		//convert.word2007ToPdf("C:\\fileupload\\convert_after\\안내문.docx", "C:\\fileupload\\convert_image", "C:\\fileupload\\convert_image\\안내문.png", "C:\\fileupload\\convert_before\\안내문.pdf");

		//PptToPng convert = new PptToPng();
		//convert.pptToPng("C:\\fileupload\\convert_after\\01장.ppt", "C:\\fileupload\\convert_image\\", "C:\\fileupload\\convert_image\\01장.png", "C:\\fileupload\\convert_before\\01장.pdf");

		//PptxToPng convert = new PptxToPng();
		//convert.pptxToPng("C:\\fileupload\\convert_after\\2021_강의자료_2장-1.pptx", "C:\\fileupload\\convert_image\\", "C:\\fileupload\\convert_image\\2021_강의자료_2장-1.png", "C:\\fileupload\\convert_before\\2021_강의자료_2장-1.pdf");

		//XlsToHtml convert = new XlsToHtml();
		//convert.xlsToHtml("C:\\fileupload\\convert_after\\대학강의시간표.xls", "C:\\fileupload\\convert_html", "C:\\fileupload\\convert_html\\대학강의시간표.html", "C:\\fileupload\\convert_after\\대학강의시간표.pdf");

		//XlsxToPDF convert = new XlsxToPDF();
		//convert.xlsxToHtml("C:\\fileupload\\convert_after\\2023학년도 1학기 생활디자인학과 시간표(안).xlsx", null, null, "C:\\fileupload\\convert_before\\2023학년도 1학기 생활디자인학과 시간표(안).pdf");
		//convert.convertExcelToPDF("C:\\fileupload\\convert_after\\2023학년도 1학기 생활디자인학과 시간표(안).xlsx", "C:\\fileupload\\convert_before\\2023학년도 1학기 생활디자인학과 시간표(안).pdf");

		XlsxToHtml convert = new XlsxToHtml();
		convert.readExcelToHtml("C:\\fileupload\\convert_after\\대학강의시간표.xls", "C:\\fileupload\\convert_html\\대학강의시간표.html", false, "xlsx", "대학강의시간표");
		//convert.readExcelToHtml("C:\\fileupload\\convert_after\\2023학년도 1학기 생활디자인학과 시간표(안).xlsx", "C:\\fileupload\\convert_html\\2023학년도 1학기 생활디자인학과 시간표(안).html", true, "xlsx", "2023학년도 1학기 생활디자인학과 시간표(안)");
	}
}
