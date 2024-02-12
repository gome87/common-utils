package com.lifeone.utils.poi2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class XlsxToPDF {

	private static final Logger LOG = LoggerFactory.getLogger(XlsxToPDF.class);

	// https://github.com/eugenp/tutorials/blob/master/text-processing-libraries-modules/pdf-2/src/main/java/com/baeldung/exceltopdf/ExcelToPDFConverter.java
	// Office(Excel) -> Html
	public void xlsxToHtml(String sOrgPath, String sImgRootPath, String sImgPath, String sChgPath) throws Exception {
		try {
			Workbook my_xls_workbook = WorkbookFactory.create(new File(sOrgPath));

			Sheet my_worksheet = my_xls_workbook.getSheetAt(0);

			short availableColumns = my_worksheet.getRow(0).getLastCellNum();

			Iterator<Row> rowIterator = my_worksheet.iterator();

			Document iText_xls_2_pdf = new Document();
			PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream(sChgPath));
			iText_xls_2_pdf.open();

			PdfPTable my_table = new PdfPTable(availableColumns);

			PdfPCell table_cell = null;

			BaseFont objBaseFont = BaseFont.createFont("font/malgun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
			Font objFont = new Font(objBaseFont);

			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch(cell.getCellType()) {
						case STRING:
							table_cell = new PdfPCell(new Phrase(cell.getStringCellValue(), objFont));
							my_table.addCell(table_cell);
							break;
						case NUMERIC:
							table_cell = new PdfPCell(new Phrase(String.valueOf(cell.getNumericCellValue()), objFont));
							my_table.addCell(table_cell);
							break;
						case BLANK:
							table_cell = new PdfPCell(new Phrase("", objFont));
							my_table.addCell(table_cell);
							break;
						default :
							try {
								table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));
							} catch (IllegalStateException illegalStateException) {
								if(illegalStateException.getMessage().equals("Cannot get a STRING value from a NUMERIC cell")) {
									table_cell = new PdfPCell(new Phrase(String.valueOf(cell.getNumericCellValue())));
								}
							}

							my_table.addCell(table_cell);
							break;
					}
				}
			}

			iText_xls_2_pdf.add(my_table);
			iText_xls_2_pdf.close();
			my_xls_workbook.close();

		} catch (Exception e) {
			LOG.error(">>>>>>>>>> convertXlsxToPdf = {}", e.getMessage());
		}
	}

	public XSSFWorkbook readExcelFile(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        inputStream.close();
        return workbook;
    }

    private Document createPDFDocument(String pdfFilePath) throws IOException, DocumentException {
        Document document = new Document();
        PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));
        document.open();
        return document;
    }

    public void convertExcelToPDF(String excelFilePath, String pdfFilePath) throws IOException, DocumentException {
        XSSFWorkbook workbook = readExcelFile(excelFilePath);
        Document document = createPDFDocument(pdfFilePath);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet worksheet = workbook.getSheetAt(i);

            // Add header with sheet name as title
            Paragraph title = new Paragraph(worksheet.getSheetName(), new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD));
            title.setSpacingAfter(20f);
            title.setAlignment(Element.ALIGN_CENTER);
            document.add(title);

            createAndAddTable(worksheet, document);
            // Add a new page for each sheet (except the last one)
            if (i < workbook.getNumberOfSheets() - 1) {
                document.newPage();
            }
        }

        document.close();
        workbook.close();
    }

    private void createAndAddTable(XSSFSheet worksheet, Document document) throws DocumentException, IOException {
        PdfPTable table = new PdfPTable(worksheet.getRow(0)
            .getPhysicalNumberOfCells());
        table.setWidthPercentage(100);
        addTableHeader(worksheet, table);
        addTableData(worksheet, table);
        document.add(table);
    }

    private void addTableHeader(XSSFSheet worksheet, PdfPTable table) throws DocumentException, IOException {
        Row headerRow = worksheet.getRow(0);
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell cell = headerRow.getCell(i);
            String headerText = getCellText(cell);
            PdfPCell headerCell = new PdfPCell(new Phrase(headerText, getCellStyle(cell)));
            setBackgroundColor(cell, headerCell);
            setCellAlignment(cell, headerCell);
            table.addCell(headerCell);
        }
    }

    public String getCellText(Cell cell) {
        String cellValue;

        try {
        	LOG.info("##### 값 확인 : {} #####" , cell.getCellType());

        	switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellValue = String.valueOf(BigDecimal.valueOf(cell.getNumericCellValue()));
                break;
            case BLANK:
            default:
                cellValue = "";
                break;
            }
		} catch (Exception e) {
			return "";
		}

        return cellValue;
    }

    private void addTableData(XSSFSheet worksheet, PdfPTable table) throws DocumentException, IOException {
        Iterator<Row> rowIterator = worksheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                continue;
            }
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                Cell cell = row.getCell(i);
                String cellValue = getCellText(cell);
                PdfPCell cellPdf = new PdfPCell(new Phrase(cellValue, getCellStyle(cell)));
                setBackgroundColor(cell, cellPdf);
                setCellAlignment(cell, cellPdf);
                table.addCell(cellPdf);
            }
        }
    }

    private void setBackgroundColor(Cell cell, PdfPCell cellPdf) {
        // Set background color

    	try {
    		short bgColorIndex = cell.getCellStyle()
    	            .getFillForegroundColor();
    	        if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
    	            XSSFColor bgColor = (XSSFColor) cell.getCellStyle()
    	                .getFillForegroundColorColor();
    	            if (bgColor != null) {
    	                byte[] rgb = bgColor.getRGB();
    	                if (rgb != null && rgb.length == 3) {
    	                    cellPdf.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
    	                }
    	            }
    	        }
		} catch (Exception e) {

		}
    }

    private void setCellAlignment(Cell cell, PdfPCell cellPdf) {

    	try {
    		CellStyle cellStyle = cell.getCellStyle();

            HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();
            VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();

            switch (horizontalAlignment) {
            case LEFT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_LEFT);
                break;
            case CENTER:
                cellPdf.setHorizontalAlignment(Element.ALIGN_CENTER);
                break;
            case JUSTIFY:
            case FILL:
                cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case RIGHT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_RIGHT);
                break;
            }

            switch (verticalAlignment) {
            case TOP:
                cellPdf.setVerticalAlignment(Element.ALIGN_TOP);
                break;
            case CENTER:
                cellPdf.setVerticalAlignment(Element.ALIGN_MIDDLE);
                break;
            case JUSTIFY:
                cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case BOTTOM:
                cellPdf.setVerticalAlignment(Element.ALIGN_BOTTOM);
                break;
            }
		} catch (Exception e) {

		}
    }

    private Font getCellStyle(Cell cell) throws DocumentException, IOException {
        Font font = new Font();

        try {
        	CellStyle cellStyle = cell.getCellStyle();
            org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet()
                .getWorkbook()
                .getFontAt(cellStyle.getFontIndexAsInt());

            short fontColorIndex = cellFont.getColor();
            if (fontColorIndex != IndexedColors.AUTOMATIC.getIndex() && cellFont instanceof XSSFFont) {
                XSSFColor fontColor = ((XSSFFont) cellFont).getXSSFColor();
                if (fontColor != null) {
                    byte[] rgb = fontColor.getRGB();
                    if (rgb != null && rgb.length == 3) {
                        font.setColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                    }
                }
            }

            if (cellFont.getItalic()) {
                font.setStyle(Font.ITALIC);
            }

            if (cellFont.getStrikeout()) {
                font.setStyle(Font.STRIKETHRU);
            }

            if (cellFont.getUnderline() == 1) {
                font.setStyle(Font.UNDERLINE);
            }

            short fontSize = cellFont.getFontHeightInPoints();
            font.setSize(fontSize);

            if (cellFont.getBold()) {
                font.setStyle(Font.BOLD);
            }

            String fontName = cellFont.getFontName();
            if (FontFactory.isRegistered(fontName)) {
                font.setFamily(fontName); // Use extracted font family if supported by iText
            } else {
                LOG.warn("Unsupported font type: {}", fontName);
                // - Use a fallback font (e.g., Helvetica)
                font.setFamily("Helvetica");
            }
		} catch (Exception e) {

		}

        return font;
    }

}
