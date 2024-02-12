package com.lifeone.utils;

import java.io.IOException;
import java.text.DateFormat;
import java.util.Date;
import java.util.Locale;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * Handles requests for the application home page.
 */
@Controller
public class HomeController {

	private static final Logger logger = LoggerFactory.getLogger(HomeController.class);

	/**
	 * Simply selects the home view to render by returning its name.
	 * @throws IOException
	 */
	@RequestMapping(value = "/", method = RequestMethod.GET)
	public String home(Locale locale, Model model) throws IOException {
		logger.info("Welcome home! The client locale is {}.", locale);

		//POIMain.convert("C:\\fileupload\\convert_after\\신청서(Application form).doc", "C:\\fileupload\\convert_before\\신청서(Application form).pdf");

		//POIMain.convert("C:\\fileupload\\convert_after\\안내문.docx", "C:\\fileupload\\convert_before\\안내문.pdf");

		//POIMain.convert("C:\\fileupload\\convert_after\\01장.ppt", "C:\\fileupload\\convert_before\\01장.pdf");

		//POIMain.convert("C:\\fileupload\\convert_after\\2021_강의자료_2장-1.pptx", "C:\\fileupload\\convert_before\\2021_강의자료_2장-1.pdf");


		Date date = new Date();
		DateFormat dateFormat = DateFormat.getDateTimeInstance(DateFormat.LONG, DateFormat.LONG, locale);

		String formattedDate = dateFormat.format(date);

		model.addAttribute("serverTime", formattedDate );

		return "home";
	}

}
