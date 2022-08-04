package com.labforward.grapecity.testgrapecity.rest;

import com.grapecity.documents.excel.DeserializationOptions;
import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

@Slf4j
@RestController
@RequestMapping("export")
public class ExportController {

	public static final String DEFAULT_OUTPUT_DIRECTORY = ".";

	@GetMapping("xlsx")
	public String saveToXlsx(@RequestParam(required = false, defaultValue = "compact-large") String fileName,
	                         @RequestParam(required = false, defaultValue = DEFAULT_OUTPUT_DIRECTORY) String outputDirectory) throws IOException {
		Workbook workbook = initializeWorkbook(fileName);
		log.debug("about to call save");
		workbook.save(String.format("%s/data-%s.xlsx", outputDirectory, System.currentTimeMillis()));
		log.debug("work done");
		return "done";
	}

	@GetMapping("pdf")
	public String saveToPdf(@RequestParam(required = false, defaultValue = "compact-large") String fileName,
	                        @RequestParam(required = false, defaultValue = DEFAULT_OUTPUT_DIRECTORY) String outputDirectory) throws IOException {
		Workbook workbook = initializeWorkbook(fileName);
		log.debug("about to call save");
		workbook.save(String.format("%s/data-%s.pdf", outputDirectory, System.currentTimeMillis()), SaveFileFormat.Pdf);
		log.debug("work done");
		return "done";
	}

	private Workbook initializeWorkbook(String fileName) throws IOException {
		String jsonData = retrieveFile(fileName);
		Workbook workbook = new Workbook();

		log.debug("about to call fromJson");
		DeserializationOptions deserializationOptions = new DeserializationOptions();
		deserializationOptions.setDoNotRecalculateAfterLoad(true);
		workbook.fromJson(jsonData, deserializationOptions);
		return workbook;
	}

	private String retrieveFile(String fileName) throws IOException {
		InputStream inputStream =
				getClass().getClassLoader().getResourceAsStream(String.format("data/%s.json", fileName));
		return new String(inputStream.readAllBytes(), StandardCharsets.UTF_8);
	}
}
