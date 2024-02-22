package com.lifeone.utils.poi;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelConverter {

	private static final Logger LOG = LoggerFactory.getLogger(ExcelConverter.class);

	// https://blog.csdn.net/qq_38567039/article/details/88418965

	/**
	 * converter excel2007(xlsx)을 HTML 변환
	 *
	 * @param String sFilePath      파일경로
	 * @param String sHtmlImageDir  HTML 이미지 경로
	 * @param String sHtmlPath      HTML 파일 경로
	 * @return String 이미지 변환 경로
	 * @throws Exception
	 * @since 2024. 02. 23
	 * @author 김영우
	 */
	public String convertXlsxToHtml(String sFilePath, String sHtmlImageDir, String sHtmlPath) throws Exception {

		InputStream is = null;
		String html = "";

		try {
			// 엑셀 읽기
			is = new FileInputStream(sFilePath);
			XSSFWorkbook workbook = new XSSFWorkbook(is);

			for(int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
				Sheet sheet = workbook.getSheetAt(numSheet);
				if(sheet == null) {
					continue;
				}

				html += "=======================" + sheet.getSheetName() + "=======================<br><br>";

				int firstRowIndex = sheet.getFirstRowNum();
				int lastRowIndex = sheet.getLastRowNum();

				html += "<table border='1' align='left'>";
				Row firstRow = sheet.getRow(firstRowIndex);
				if(firstRow == null) {
					continue;
				}
				for(int i =firstRow.getFirstCellNum(); i <= firstRow.getLastCellNum(); i++) {
					Cell cell = firstRow.getCell(i);
					String cellValue = this.getCellValue(cell, true);
					html += "<th>" + cellValue + "</th>";
				}

				// 행
				for(int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++) {
					Row currentRow = sheet.getRow(rowIndex);
					html += "<tr>";
					if(currentRow != null) {
						int firstColumnIndex = currentRow.getFirstCellNum();
						int lastColumnIndex = currentRow.getLastCellNum();

						// 열
						for(int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
							Cell currentCell = currentRow.getCell(columnIndex);
							String currentCellValue = this.getCellValue(currentCell, true);
							html += "<td>" + currentCellValue + "</td>";
						}
					} else {
						html += " ";
					}
					html += "</tr>";
				}
				html += "</table>";


				ByteArrayOutputStream outStream = new ByteArrayOutputStream();
				DOMSource domSource = new DOMSource();
				StreamResult streamResult = new StreamResult(outStream);

				TransformerFactory tf = TransformerFactory.newInstance();
				Transformer serializer = tf.newTransformer();
				serializer.setOutputProperty(OutputKeys.ENCODING, "gbk");
				serializer.setOutputProperty(OutputKeys.INDENT, "yes");
				serializer.setOutputProperty(OutputKeys.METHOD, "html");

				serializer.transform(domSource, streamResult);
				outStream.close();

				String sContent = new String(outStream.toByteArray());

				// HTML 파일 생성
				FileUtils.writeStringToFile(new File(sHtmlPath), sContent, "gbk");
			}
		} catch (Exception e) {
			LOG.error(">>>>>>>>>> convertXlsxToHtml = {}", e.getMessage());
			throw new Exception("##### convertXlsxToHtml Error 발생 #####");
		}

		return sHtmlPath;
	}

	private String getCellValue(Cell cell, boolean treatAsStr) {
		if(cell == null) {
			return "";
		}
		if(treatAsStr) {
			cell.getCellType();
		}

		if(cell.getCellType() == CellType.BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());
		} else if(cell.getCellType() == CellType.NUMERIC) {
			return String.valueOf(cell.getNumericCellValue());
		} else {
			return String.valueOf(cell.getStringCellValue());
		}
	}
}
