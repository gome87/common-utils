package com.lifeone.utils.poi2;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class XlsxToHtml {

	private static final Logger LOG = LoggerFactory.getLogger(XlsxToHtml.class);

	// https://blog.csdn.net/qq_33697094/article/details/122736603
	public String readExcelToHtml(String filePath, String htmlPosition, boolean isWithStyle, String type, String attname) throws Exception {

		InputStream is = null;
		String htmlExcel = null;
		Map<String, String> stylemap = new HashMap<>();

		try {
			if("csv".equalsIgnoreCase(type)) {
				htmlExcel = this.getCSVInfo(filePath, htmlPosition);
				this.writeFile1(htmlExcel, htmlPosition, stylemap, attname);
			} else {
				File sourceFile = new File(filePath);
				is = new FileInputStream(sourceFile);
				Workbook wb = WorkbookFactory.create(sourceFile);
				if(wb instanceof XSSFWorkbook) {
					XSSFWorkbook xWb = (XSSFWorkbook) wb;
					htmlExcel = this.getExcelInfo(xWb, isWithStyle, stylemap);
				} else if (wb instanceof HSSFWorkbook) {
					HSSFWorkbook hWb = (HSSFWorkbook) wb;
					htmlExcel = this.getExcelInfo(hWb, isWithStyle, stylemap);
				}
				writeFile(htmlExcel, htmlPosition, stylemap, attname);
			}
		} catch (Exception e) {
			LOG.error(">>>>>>>>>> readExcelToHtml = {}", e.getMessage());
		} finally {
			try {
				if(is != null) {
					is.close();
				}
			} catch (IOException e2) {
				LOG.error(">>>>>>>>>> readExcelToHtml Error = {}", e2.getMessage());
			}
		}
		return htmlPosition;
	}

	private void getcsvvalue(BufferedReader reader, List<String> col, String oldvalue, List<List<String>> list) {
		String line  = null;
		try {
			while((line=reader.readLine()) != null) {
				String[] item = line.split(",", -1);
				boolean isbreak = false;
				for(int i=0; i<item.length; i++) {
					String value = item[i];
					if(value.endsWith("\"")) {
						value = oldvalue + value;
						col.add(value);
					} else if(item.length == 1) {
						value = oldvalue + value;
						this.getcsvvalue(reader, col, value, list);
						isbreak = true;
					} else if(value.startsWith("\"")) {
						this.getcsvvalue(reader, col, value, list);
						isbreak = true;
					} else {
						col.add(value);
					}
				}

				if(!isbreak) {
					list.add(col);
					col = new ArrayList();
				}
			}
		} catch (IOException e) {
			LOG.error(">>>>>>>>>> getcsvvalue = {}", e.getMessage());
		}
	}

	private String getCSVInfo(String filePath, String htmlPositon) {
		StringBuffer sb = new StringBuffer();
		DataInputStream in  = null;
		try {
			in = new DataInputStream(new FileInputStream(filePath));
			BufferedReader reader = new BufferedReader(new InputStreamReader(in));

			String  line  = null;
			List<List<String>> list = new ArrayList<>();
			while((line = reader.readLine()) != null) {
				String[] item = line.split(",");
				List<String> col = new ArrayList<>();
				for(int i=0; i<item.length; i++) {
					String value = item[i];
					if(value.startsWith("\"")) {
						this.getcsvvalue(reader, col, value, list);
					} else {
						col.add(value);
					}
				}
				list.add(col);
			}

			sb.append("<table>");
			for(int i=0; i<list.size(); i++) {
				List<String> col = (List)list.get(i);
				if(col == null || col.size() == 0) {
					sb.append("<tr><td></td></tr>");
				}
				sb.append("<tr>");
				for(int j=0;j<col.size(); j++) {
					String value = (String) col.get(j);
					if(value == null || "".equals(value)) {
						sb.append("<td></td>");
						continue;
					} else {
						sb.append("<td>"+value+"</td>");
					}
				}
				sb.append("</tr>");
			}
			sb.append("</table>");
		} catch (IOException e) {
			LOG.error(">>>>>>>>>> getCSVInfo = {}", e.getMessage());
		} finally {
			try {
				in.close();
			} catch (IOException e2) {
				LOG.error(">>>>>>>>>> getCSVInfo Error = {}", e2.getMessage());
			}
		}
		return sb.toString();
	}

	private String getExcelInfo(Workbook wb, boolean isWithStyle, Map<String, String> stylemap) {

		StringBuffer sb = new StringBuffer();
		StringBuffer ulsb = new StringBuffer();

		ulsb.append("<ul>");
		int num = wb.getNumberOfSheets();

		for(int i=0; i<num; i++) {
			Sheet sheet = wb.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			if(i==0) {
				ulsb.append("<li id='li_" + i + "' class='cur' onclick='changetab(" + i + ")'>" + sheetName + "</li>");
			} else {
				ulsb.append("<li id='li_" + i + "' onclick='changettab(" + i + ")'>" + sheetName + "</li>");
			}
			int lastRowNum = sheet.getLastRowNum();
			Map<String, String> map[] = this.getRowSpanColSpanMap(sheet);
			Map<String, String> map1[] = this.getRowSpanColSpanMap(sheet);
			sb.append("<table id='table_" + i + "' ");
			if(i == 0) {
				sb.append("class='block'");
			}
			sb.append(">");

			Row row = null;
			Cell cell = null;

			int maxRowNum = 0;
			int maxColNum = 0;

			for(int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
				row = sheet.getRow(rowNum);
				if(row == null) {
					continue;
				}
				int lastColNum = row.getLastCellNum();
				for(int colNum = 0; colNum < lastColNum; colNum++) {
					cell = row.getCell(colNum);
					if(cell == null) {
						continue;
					}
					String stringValue = this.getCellValue1(cell);
					if(map1[0].containsKey(rowNum + "," + colNum)) {
						map1[0].remove(rowNum + "," + colNum);
						if(maxRowNum < rowNum) {
							maxRowNum = rowNum;
						}
						if(maxColNum < colNum) {
							maxColNum = colNum;
						}
					} else if(map1[1].containsKey(rowNum + "," + colNum)) {
						map1[1].remove(rowNum + "," + colNum);
						if(maxRowNum < rowNum) {
							maxRowNum = rowNum;
						}
						if(maxColNum < colNum) {
							maxColNum = colNum;
						}
						continue;
					}

					if(stringValue == null || "".equals(stringValue.trim())) {
						continue;
					} else {
						if(maxRowNum < rowNum) {
							maxRowNum = rowNum;
						}
						if(maxColNum < colNum) {
							maxColNum = colNum;
						}
					}
				}
			}

			for(int rowNum = sheet.getFirstRowNum(); rowNum <= maxRowNum; rowNum++) {
				row = sheet.getRow(rowNum);
				if(row == null) {
					sb.append("<tr><td></td></td>");
					continue;
				}
				sb.append("<tr>");

				int lastColNum = row.getLastCellNum();
				for(int colNum = 0; colNum <= maxColNum; colNum++) {
					cell = row.getCell(colNum);
					if(cell == null) {
						sb.append("<td></td>");
						continue;
					}
					String stringValue = this.getCellValue(cell);
					if(map[0].containsKey(rowNum + "," + colNum)) {
						String pointString = map[0].get(rowNum + "," + colNum);
						map[0].remove(rowNum + "," + colNum);
						int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
						int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
						int rowSpan = bottomeRow - rowNum + 1;
						int colSpan = bottomeCol - colNum + 1;
						sb.append("<td rowspan = '" + rowSpan + "' colspan= '" + colSpan + "' ");
					} else if(map[1].containsKey(rowNum + "," + colNum)) {
						map[1].remove(rowNum + "," + colNum);
						continue;
					} else {
						sb.append("<td");
					}

					if(isWithStyle) {
						this.dealExcelStyle(wb, sheet, cell, sb, stylemap);
					}

					sb.append("><nobr>");

					if(stringValue == null || "".equals(stringValue.trim())) {
						FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
						if(evaluator.evaluate(cell) != null) {
							String cellnumber = evaluator.evaluate(cell).getNumberValue() + "";

							if(null != cellnumber && cellnumber.contains(".")) {
								String[] decimal = cellnumber.split("\\.");
								if(decimal[1].length() > 2) {
									int num1 = decimal[1].charAt(0) - '0';
									int num2 = decimal[1].charAt(1) - '0';
									int num3 = decimal[1].charAt(2) - '0';
									if(num3 == 9) {
										num2 = 0;
									} else if(num3 >= 5) {
										num2 = num2 + 1;
									}
									cellnumber = decimal[0] + "." + num1 + num2;
								}
							}
							stringValue = cellnumber;
						}
						sb.append(stringValue.replace(String.valueOf((char) 160), " "));
					} else {
						sb.append(stringValue.replace(String.valueOf((char) 160), " "));
					}
					sb.append("</nobr></td>");
				}
				sb.append("</tr>");
			}
			sb.append("</table>");
		}
		ulsb.append("</ul>");

		return ulsb+toString() + sb.toString();
	}

	private Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {
		Map<String, String> map0 = new HashMap<>();
		Map<String, String> map1 = new HashMap<>();
		int mergedNum = sheet.getNumMergedRegions();
		CellRangeAddress range = null;
		for(int i = 0; i < mergedNum; i++) {
			range = sheet.getMergedRegion(i);
			int topRow = range.getFirstRow();
			int topCol = range.getFirstColumn();
			int bottomRow = range.getLastRow();
			int bottomCol = range.getLastColumn();
			map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);

			int tempRow = topRow;
			while(tempRow <= bottomRow) {
				int tempCol = topCol;
				while(tempCol <= bottomCol) {
					map1.put(tempRow + "," + tempCol, "");
					tempCol++;
				}
				tempRow++;
			}
			map1.remove(topRow + "," + topCol);
		}
		Map[] map = {map0, map1};
		return map;
	}

	private String getCellValue1(Cell cell) {
		String result = new String();
		switch (cell.getCellType()) {
			case NUMERIC:
				result = "1";
				break;

			case STRING:
				result = "1";
				break;

			case BLANK:
				result = "";
				break;

			default:
				result = "";
				break;
		}
		return result;
	}

	private String getCellValue(Cell cell) {
		String result = new String();
		switch (cell.getCellType()) {
			case NUMERIC:
				if(DateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat sdf = null;
					if(cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
						sdf = new SimpleDateFormat("HH:mm");
					} else {
						sdf = new SimpleDateFormat("yyyy-MM-dd");
					}
					Date date = cell.getDateCellValue();
					result = sdf.format(date);
				} else if(cell.getCellStyle().getDataFormat() == 58) {
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					double value = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(value);
					result = sdf.format(date);
				} else {
					double value = cell.getNumericCellValue();
					CellStyle style = cell.getCellStyle();
					DecimalFormat format = new DecimalFormat();
					String temp = style.getDataFormatString();

					if(temp.equals("General")) {
						format.applyPattern("#");
					}
					result = format.format(value);
				}
				break;
			case STRING:
				result = cell.getRichStringCellValue().toString();
				break;
			case BLANK:
				result = "";
				break;
			default:
				result = "";
				break;
		}
		return result;
	}

	private void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb, Map<String, String> stylemap) {
		CellStyle cellStyle = cell.getCellStyle();
		if(cellStyle != null) {
			HorizontalAlignment alignment = cellStyle.getAlignment();

			VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
			String _style = "vertical-align:" + this.convertVerticalAlignToHtml(verticalAlignment) + ";";
			if(wb instanceof XSSFWorkbook) {
				XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();

				short boldWeight = 400;
				String align = this.convertAlignToHtml(alignment);
				int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
				_style += "font-weight:" + boldWeight + ";font-size: " + xf.getFontHeight() / 2 + "%;width:" + columnWidth + "px;text-align:" + align + ";";

				XSSFColor xc = xf.getXSSFColor();
				if(xc != null && !"".equals(xc)) {
					_style += "color:#" + xc.getARGBHex().substring(2) + ";";
				}

				XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
				if(bgColor != null && !"".equals(bgColor)) {
					_style += "background-color:#" + bgColor.getARGBHex().substring(2) + ";";
				}
				_style += this.getBorderStyle(0, cellStyle.getBorderTop().getCode(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor());
				_style += this.getBorderStyle(1, cellStyle.getBorderRight().getCode(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor());
				_style += this.getBorderStyle(3, cellStyle.getBorderLeft().getCode(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor());
				_style += this.getBorderStyle(2, cellStyle.getBorderBottom().getCode(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor());
			} else if(wb instanceof HSSFWorkbook) {
				HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
				short boldWeight = hf.getFontHeight();
				short fontColor = hf.getColor();
				HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette();
				HSSFColor hc = palette.getColor(fontColor);
				String align = this.convertAlignToHtml(alignment);
				int columnWidht = sheet.getColumnWidth(cell.getColumnIndex());
				_style += "font-weight:" + boldWeight + ";font-size: " + hf.getFontHeight() / 2 + "%;text-align:" + align + ";width:" + columnWidht + "px;";
				String fontColorStr = this.convertToStardColor(hc);
				if(fontColorStr != null && !"".equals(fontColorStr.trim())) {
					_style += "color:" + fontColorStr + ";";
				}
				short bgColor = cellStyle.getFillForegroundColor();
				hc = palette.getColor(bgColor);
				String bgColorStr = this.convertToStardColor(hc);
				if(bgColorStr != null && !"".equals(bgColorStr.trim())) {
					_style += "background-color:" + bgColorStr + ";";
				}
				_style += this.getBorderStyle(palette, 0, cellStyle.getBorderTop().getCode(), cellStyle.getTopBorderColor());
				_style += this.getBorderStyle(palette, 1, cellStyle.getBorderRight().getCode(), cellStyle.getRightBorderColor());
				_style += this.getBorderStyle(palette, 3, cellStyle.getBorderLeft().getCode(), cellStyle.getLeftBorderColor());
				_style += this.getBorderStyle(palette, 2, cellStyle.getBorderBottom().getCode(), cellStyle.getBottomBorderColor());
			}
			String calssname = "";
			if(!stylemap.containsKey(_style)) {
				int count = stylemap.size();
				calssname = "td" + count;
				stylemap.put(_style, calssname);
			} else {
				calssname = stylemap.get(_style);
			}
			if(!"".equals(calssname)) {
				sb.append("class='"+ calssname + "'");
			}
		}
	}

	private String convertAlignToHtml(HorizontalAlignment alignment) {
		String align = "center";
		switch (alignment) {
			case LEFT:
				align = "left";
				break;
			case CENTER:
				align = "center";
				break;
			case RIGHT:
				align = "right";
				break;
			default:
				break;
		}
		return align;
	}

	private String convertVerticalAlignToHtml(VerticalAlignment verticalAlignment) {
		String valign = "middle";
		switch (verticalAlignment) {
			case BOTTOM:
				valign = "bottom";
				break;
			case CENTER:
				valign = "middle";
				break;
			case TOP:
				valign = "top";
				break;
			default:
				break;
		}
		return valign;
	}

	private String convertToStardColor(HSSFColor hc) {
		StringBuffer sb = new StringBuffer("");
		if(hc != null) {
			if(HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex() == hc.getIndex()) {
				return null;
			}
			sb.append("#");
			for(int i = 0; i < hc.getTriplet().length; i++) {
				sb.append(this.fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
			}
		}
		return sb.toString();
	}

	private String fillWithZero(String str) {
		if(str != null && str.length() < 2) {
			return "0" + str;
		}
		return str;
	}

	static String[] bordesr = {"border-top:", "border-right:", "border-bottom:", "border-left:"};
	static String[] borderStyles = {"solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid "};

	private String getBorderStyle(HSSFPalette palette, int b, short s, short t) {
		if(s == 0) {
			return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
		}
		String borderColorStr = this.convertToStardColor(palette.getColor(t));
		borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
		return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
	}

	private String getBorderStyle(int b, short s, XSSFColor xc) {
		if(s == 0) {
			return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
		}
		if(xc != null && !"".equals(xc)) {
			String borderColorStr = xc.getARGBHex();
			borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr.substring(2);
			return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
		}
		return "";
	}

	private void writeFile(String content, String htmlPath, Map<String, String> stylemap, String name) {

		File file2 = new File(htmlPath);
		StringBuffer sb = new StringBuffer();

		try {
			file2.createNewFile();
			sb.append("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><title>" + name + "</title><style type=\"text/css\">");
			sb.append("ul{list-style; none;max-widht: calc(100%);padding: 0px;margin: 0px;overflow-x: scroll;white-space: nowrap;} ul li{padding: 3px 5px;display: inline-block;border-right: 1px solid #768893;} ul li.cur{color: #F59C25;} table{border-collapse: collapse;display: none;width: 100%;} table.block{display: block;}");
			for(Map.Entry<String, String> entry : stylemap.entrySet()) {
				String mapKey = entry.getKey();
				String mapValue = entry.getValue();
				sb.append(" ." + mapValue + "{" + mapKey + "}");
			}
			sb.append("</style><script>");
			sb.append("function changetab(i){var block = document.getElementsByClassName(\"block\");block[0].className = block[0].className.replace(\"block\",\"\");var cur = document.getElementsByClassName(\"cur\");cur[0].className = cur[0].className.replace(\"cur\",\"\");var curli = document.getElementById(\"li_\"+i);curli.className += ' cur';var curtable = document.getElementById(\"table_\"+i);curtable.className=' block';}");
			sb.append("</script></head><body>");
			sb.append("<div>");
			sb.append(content);
			sb.append("</div>");
			sb.append("</body></html>");
			FileUtils.write(file2, sb.toString(), "UTF-8");
		} catch (Exception e) {
			LOG.error(">>>>>>>>>> writeFile Error = {}", e.getMessage());
		}
	}

	private void writeFile1(String content, String htmlPath, Map<String, String> stylemap, String name) {
		File file2 = new File(htmlPath);
		StringBuffer sb = new StringBuffer();
		try {
			file2.createNewFile();
			sb.append("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><title>" + name + "</title><style type=\"text/css\">");
			sb.append("ul{list-style; none;max-widht: calc(100%);padding: 0px;margin: 0px;overflow-x: scroll;white-space: nowrap;} ul li{padding: 3px 5px;display: inline-block;border-right: 1px solid #768893;} ul li.cur{color: #F59C25;} table{border-collapse: collapse;width: 100%;} td{border: solid #000000 1px; min-width: 200px;}");
			sb.append("</style></head><body>");
			sb.append("<div>");
			sb.append(content);
			sb.append("</div>");
			sb.append("</body></html>");
			FileUtils.write(file2, sb.toString(), "UTF-8");
		} catch (Exception e) {
			LOG.error(">>>>>>>>>> writeFile Error = {}", e.getMessage());
		}
	}

}
