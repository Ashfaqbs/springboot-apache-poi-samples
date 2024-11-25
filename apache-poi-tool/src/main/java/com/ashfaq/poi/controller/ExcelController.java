package com.ashfaq.poi.controller;

import java.io.ByteArrayOutputStream;
import java.util.List;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.ashfaq.poi.service.ExcelService;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

	private final ExcelService excelService;

	public ExcelController(ExcelService excelService) {
		this.excelService = excelService;
	}

	@PostMapping("/compare-columns")
	public ResponseEntity<List<String>> compareColumns(@RequestParam("file") MultipartFile file) {
		try {
			List<String> matches = excelService.findSimilarWords(file);
			return ResponseEntity.ok(matches);
		} catch (Exception e) {
			return ResponseEntity.badRequest().body(List.of("Error: " + e.getMessage()));
		}
	}

	@PostMapping("/download-matches")
	public ResponseEntity<byte[]> downloadMatchingWords(@RequestParam("file") MultipartFile file) {
		try {
			// Generate the Excel file with matching words
			ByteArrayOutputStream outputStream = excelService.generateExcelWithMatches(file);

			// Prepare the response
			return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=common_words.xlsx")
					.contentType(MediaType.APPLICATION_OCTET_STREAM).body(outputStream.toByteArray());
		} catch (Exception e) {
			return ResponseEntity.badRequest().body(("Error: " + e.getMessage()).getBytes());
		}
	}

}
