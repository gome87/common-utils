package com.lifeone.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@Controller
public class OfficeController {

	private static final Logger LOG = LoggerFactory.getLogger(OfficeController.class);

	/**
	 * Office -> HTML 테스트용
	 *
	 * @param HttpServletRequest  req
	 * @param HttpServletResponse rep
	 * @return
	 * @throws Exception
	 * @since 2024. 02. 23
	 * @author 김영우
	 */
	@RequestMapping(value = "/office", method = { RequestMethod.GET })
	public void retrieveOfficeChange(HttpServletRequest req, HttpServletResponse res) throws Exception {


	}

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

		//XlsxToHtml convert = new XlsxToHtml();
		//convert.readExcelToHtml("C:\\fileupload\\convert_after\\대학강의시간표.xls", "C:\\fileupload\\convert_html\\대학강의시간표.html", false, "xlsx", "대학강의시간표");
		//convert.readExcelToHtml("C:\\fileupload\\convert_after\\2023학년도 1학기 생활디자인학과 시간표(안).xlsx", "C:\\fileupload\\convert_html\\2023학년도 1학기 생활디자인학과 시간표(안).html", true, "xlsx", "2023학년도 1학기 생활디자인학과 시간표(안)");
	}

	/**
	 * filePath search
	 *
	 * @param HttpServletRequest  req
	 * @param HttpServletResponse rep
	 * @return
	 * @throws Exception
	 * @since 2024. 02. 23
	 * @author 김영우
	 */
	@RequestMapping(value = "/filePathSearch", method = { RequestMethod.GET })
	public void retrieveFilePathSearch(HttpServletRequest req, HttpServletResponse res) throws Exception {
		List<String> fileLst = new ArrayList<> ();

		this.scanDir("C:\\fileupload", fileLst);	// 테스트용 임시 폴더

		for(String fullPath : fileLst) {
			LOG.info("##### File List : {} #####", fullPath);
		}
	}

	// 재귀 호출을 이용하여 하위 폴더 탐색
	private void scanDir(String sFolderPath, List<String> fileList) {

		File[] files = new File(sFolderPath).listFiles();

		for(File fileElement : files) {
			if(fileElement.isDirectory()) {
				scanDir(fileElement.getAbsolutePath(), fileList);
			} else {
				fileList.add(fileElement.getAbsolutePath());
			}
		}
	}

}
