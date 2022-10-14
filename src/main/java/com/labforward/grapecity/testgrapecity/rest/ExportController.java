package com.labforward.grapecity.testgrapecity.rest;

import com.grapecity.documents.excel.DeserializationOptions;
import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.IWorksheets;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;


import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.BorderLineStyle;


import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ITableStyleElement;

import com.grapecity.documents.excel.DocumentProperties;
import com.grapecity.documents.excel.PdfSaveOptions;

import java.util.Calendar;
import java.util.Date;
import java.text.SimpleDateFormat;

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
		IWorksheets sheets = workbook.getWorksheets();

		Workbook.FontsFolderPath = "./src/main/resources/fonts";

		sheets.forEach(sheet -> {
			sheet.getSheetView().setGridlineColor(Color.GetSilver());
			sheet.getPageSetup().setPrintHeadings(true);
			sheet.getPageSetup().setPrintGridlines(true);
		});

		String mydate = java.text.DateFormat.getDateTimeInstance().format(Calendar.getInstance().getTime());

		Calendar today = Calendar.getInstance();
		today.clear(Calendar.HOUR); today.clear(Calendar.MINUTE); today.clear(Calendar.SECOND);
		Date todayDate = today.getTime();

		DocumentProperties documentProperties = new DocumentProperties();
		//Sets the name of the person that created the PDF document.
		documentProperties.setAuthor("Clark Williams");
		//Sets the title of thePDF document.
		documentProperties.setTitle("My Table title");
		//Set the PDF version.
		documentProperties.setPdfVersion(1.5f); // needed for what?
		//Set the subject of the PDF document.
		documentProperties.setSubject("Entry title");
		//Set the keyword associated with the PDF document.
		documentProperties.setKeywords("project title");
		//Set the creation date and time of the PDF document.
		documentProperties.setCreationDate(today);  // table last modified date?, today?
		//Set the date and time the PDF document was most recently modified.
		documentProperties.setModifyDate(today); 
		//Set the name of the application that created the original PDF document.
		documentProperties.setCreator("Labfolder");
		//Set the name of the application that created the PDF document.
		documentProperties.setProducer("Labforward");


		PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
		//Sets the document properties of the pdf.
		pdfSaveOptions.setDocumentProperties(documentProperties);

		log.debug("about to call save");
		workbook.save(String.format("%s/data-%s.pdf", outputDirectory, System.currentTimeMillis()), pdfSaveOptions);
		log.debug("work done");
		return "done";
	}

	private Workbook initializeWorkbook(String fileName) throws IOException {
		String jsonData = retrieveFile(fileName);

		// non-free font replacement with a metrically compatible replacement
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Calibri"), "\"$1\":\"$2Carlito");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Arial"), "\"$1\":\"$2Arimo");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("(\\\\\")?Times New Roman(\\\\\")?"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("(\\\\\")?Courier New(\\\\\")?"), "\"$1\":\"$2CourierPrime");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("(\\\\\")?Comic Sans MS(\\\\\")?"), "\"$1\":\"$2ComicNeue");
		
		// just used working sans or serifs
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Georgia"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Tahoma"), "\"$1\":\"$2Carlito");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Times"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Verdana"), "\"$1\":\"$2Carlito");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Cambria"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Candara"), "\"$1\":\"$2Carlito");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Century"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("Garamond"), "\"$1\":\"$2Tinos");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("(\\\\\")?Trebuchet MS(\\\\\")?"), "\"$1\":\"$2Carlito");
		jsonData = jsonData.replaceAll( fontReplacementRegexp("(\\\\\")?Arial Black(\\\\\")?"), "\"$1\":\"$2Arimo");
	

		Workbook workbook = new Workbook();

		log.debug("about to call fromJson");
		DeserializationOptions deserializationOptions = new DeserializationOptions();
		deserializationOptions.setDoNotRecalculateAfterLoad(true);
		workbook.fromJson(jsonData, deserializationOptions);
		return workbook;
	}

  private String fontReplacementRegexp(String searchFont) {
		// searches for: 
		// "font": "normal normal 14.7px Calibri",
		// "fontForColumnWidth": "normal normal 14.7px \"Courier New\""
		// "headingFont": "Cambria",
		// "bodyFont": "Calibri"

	String fontKeyword = "\"(font|fontForColumnWidth|headingFont|bodyFont)\"";
		String fontStyle = "([^,\\\\]*)";

		return fontKeyword + ":\\s*\"" + fontStyle + searchFont;
	}

	private String retrieveFile(String fileName) throws IOException {
		InputStream inputStream =
				getClass().getClassLoader().getResourceAsStream(String.format("data/%s.json", fileName));
		return new String(inputStream.readAllBytes(), StandardCharsets.UTF_8);
	}
}
