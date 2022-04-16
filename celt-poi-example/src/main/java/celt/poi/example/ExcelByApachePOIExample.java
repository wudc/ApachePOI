package celt.poi.example;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelByApachePOIExample {
	static Cell lastMonthDateCell;

	/**
	 * Create main title row
	 * 
	 * @param sheet
	 */
	static void createTitle(Sheet sheet, XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		sheet.addMergedRegion(CellRangeAddress.valueOf("A1:J1"));
		Row row = sheet.createRow(0);
		style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(font);
		Cell cell = row.createCell(0);
		cell.setCellValue("Employee Data");
		cell.setCellStyle(style);

		createSubTitle(sheet, font, style);
	}

	static void createSubTitle(Sheet sheet, Font font, CellStyle style) {
		Row row = sheet.createRow(2);
		style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(font);

		sheet.addMergedRegion(CellRangeAddress.valueOf("A3:J3"));
		Cell cell = row.createCell(0);
		cell.setCellValue("Employee Information");
		cell.setCellStyle(style);
	}

	static void createMetaData(Sheet sheet, XSSFWorkbook workbook) {
		// CellStyle style = workbook.createCellStyle();
		sheet.addMergedRegion(CellRangeAddress.valueOf("A4:B4"));
		Row row = sheet.createRow(3);
		Cell cell = row.createCell(0);
		cell.setCellValue("Company:");
		cell = row.createCell(2);
		sheet.addMergedRegion(CellRangeAddress.valueOf("C4:E4"));
		cell.setCellValue("CELT Management Inc.");

		sheet.addMergedRegion(CellRangeAddress.valueOf("A5:B5"));
		row = sheet.createRow(4);
		cell = row.createCell(0);
		cell.setCellValue("Department:");
		cell = row.createCell(2);
		sheet.addMergedRegion(CellRangeAddress.valueOf("C5:E5"));
		cell.setCellValue("Accounting and Tax");

		sheet.addMergedRegion(CellRangeAddress.valueOf("A6:B6"));
		row = sheet.createRow(5);
		cell = row.createCell(0);
		cell.setCellValue("Department Head:");
		cell = row.createCell(2);
		sheet.addMergedRegion(CellRangeAddress.valueOf("C6:E6"));
		cell.setCellValue("Director of Accounting");

		sheet.addMergedRegion(CellRangeAddress.valueOf("A7:B7"));
		row = sheet.createRow(6);
		cell = row.createCell(0);
		cell.setCellValue("Creation Date:");
		cell = row.createCell(2);
		sheet.addMergedRegion(CellRangeAddress.valueOf("C7:E7"));
		cell.setCellValue("03/08/1972");

	}

	static List<String> createSampleHoursList(int size) {
		List<String> hours = new ArrayList<>();
		for (int i = 0; i < size; i++) {
			hours.add("" + i + " AM");
		}

		return hours;
	}

	static void createHoursColumnLabel(Sheet sheet, XSSFWorkbook workbook, List<String> hours) {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		style.setFont(font);
		int initRow = 10;
		for (String hour : hours) {
			Row row = sheet.createRow(initRow++);
			Cell cell = row.createCell(0);
			cell.setCellValue(hour);
			cell.setCellStyle(style);
		}

	}

	static List<String> createSampleDate(int size, String month) {
		List<String> monthDate = new ArrayList<>();
		for (int i = 0; i < size; i++) {
			monthDate.add(month + "/" + (i + 1));
		}

		return monthDate;
	}

	static void createDataTableColumnNames(Sheet sheet, XSSFWorkbook workbook, List<String> monthDates) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		Font font = workbook.createFont();
		style.setFont(font);
		font.setBold(true);
		int initCol = 1;
		Row row = sheet.createRow(9);
		for (String monthDate : monthDates) {
			Cell cell = row.createCell(initCol++);
			cell.setCellValue(monthDate);
			cell.setCellStyle(style);
			lastMonthDateCell = cell;
		}
	}

	static void createTotalRow(Sheet sheet, XSSFWorkbook workbook, int rowNum) {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		style.setFont(font);
		font.setBold(true);
		int initCol = 1;
		Row row = sheet.createRow(rowNum);

		Cell cell = row.createCell(0);
		cell.setCellValue("Total");
		cell.setCellStyle(style);

		cell = row.createCell(1);
		cell.setCellValue("5");
		cell.setCellStyle(style);

		cell = row.createCell(2);
		cell.setCellValue("50");
		cell.setCellStyle(style);

	}

	/**
	 * Color Alternate Rows in Different Colors
	 * 
	 * @param sheet
	 */
	static void shadeAlt(Sheet sheet, String start, String end) {
		SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		// Condition 1: Formula Is =A2=A1 (White Font)
		ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
		PatternFormatting fill1 = rule1.createPatternFormatting();
		fill1.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
		fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] regions = { CellRangeAddress.valueOf(start + ":" + end) };

		sheetCF.addConditionalFormatting(regions, rule1);
	}

	static String findLastColumnName() {
		return CellReference.convertNumToColString(lastMonthDateCell.getColumnIndex());
	}

	public static void main(String[] args) {
		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		createTitle(sheet, workbook);
		createMetaData(sheet, workbook);

		List<String> hours = createSampleHoursList(5);
		createHoursColumnLabel(sheet, workbook, hours);
		List<String> monthDates = createSampleDate(6, "3");
		createDataTableColumnNames(sheet, workbook, monthDates);

		int dataTableInitRow = 10; // zero based
		String start = "A" + (dataTableInitRow + 1);
		String end = findLastColumnName() + (hours.size() + dataTableInitRow);

		int startTotalRow = hours.size() + dataTableInitRow;
		createTotalRow(sheet, workbook, startTotalRow);
		// set region with zebra look
		shadeAlt(sheet, start, end);

		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("ExcelByApachePOIExample.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("ExcelByApachePOIExample.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
