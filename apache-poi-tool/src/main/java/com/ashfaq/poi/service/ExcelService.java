package com.ashfaq.poi.service;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelService {

	public List<String> findSimilarWords(MultipartFile file) throws Exception {
		List<String> matches = new ArrayList<>();

		try (InputStream inputStream = file.getInputStream()) {
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0); // First sheet
			List<String> column1 = new ArrayList<>();
			List<String> column2 = new ArrayList<>();

			for (Row row : sheet) {
				Cell cell1 = row.getCell(0);
				Cell cell2 = row.getCell(1);

				if (cell1 != null && cell2 != null) {
					column1.add(cell1.getStringCellValue());
					column2.add(cell2.getStringCellValue());
				}
			}

			for (int i = 0; i < column1.size(); i++) {
				String word1 = column1.get(i);
				String word2 = column2.get(i); // Compare corresponding rows

				if (isSimilar(word1, word2)) {
					matches.add(word1 + " - " + word2);
				}
			}
		}

		return matches;
	}

	// Simple similarity check (can be extended with better algorithms)
	private boolean isSimilar(String word1, String word2) {
		word1 = word1.toLowerCase();
		word2 = word2.toLowerCase();
		return word1.equals(word2) || word2.contains(word1) || word1.contains(word2);
	}

	public ByteArrayOutputStream generateExcelWithMatches(MultipartFile file) throws Exception {
		List<String[]> matches = new ArrayList<>();

		// Read Excel file and find matches
		try (InputStream inputStream = file.getInputStream()) {
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0); // First sheet

			List<String> column1 = new ArrayList<>();
			List<String> column2 = new ArrayList<>();

			for (Row row : sheet) {
				Cell cell1 = row.getCell(0);
				Cell cell2 = row.getCell(1);

				if (cell1 != null && cell2 != null) {
					column1.add(cell1.getStringCellValue());
					column2.add(cell2.getStringCellValue());
				}
			}

			// Find matches between the two columns
			for (int i = 0; i < column1.size(); i++) {
				String word1 = column1.get(i);
				String word2 = column2.get(i);
				if (isSimilar(word1, word2)) {
					matches.add(new String[] { word1, word2 });
				}
			}
		}

		// Generate new Excel file with matches
		return createExcel(matches);
	}

	private ByteArrayOutputStream createExcel(List<String[]> matches) throws Exception {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Common Words");

		// Add header
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Column A");
		header.createCell(1).setCellValue("Column B");

		// Add data rows
		int rowNum = 1;
		for (String[] match : matches) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(match[0]);
			row.createCell(1).setCellValue(match[1]);
		}

		// Write to output stream
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		workbook.close();
		return outputStream;
	}
}
