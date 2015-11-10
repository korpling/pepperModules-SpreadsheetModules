package org.corpus_tools.peppermodules.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.SToken;
import org.eclipse.emf.common.util.URI;

/**
 * 
 * @author Vivian Voigt
 *
 */
public class Spreadsheet2SaltMapper extends PepperMapperImpl implements
		PepperMapper {

	public SpreadsheetImporterProperties getProps() {
		return ((SpreadsheetImporterProperties) this.getProperties());
	}

	@Override
	public DOCUMENT_STATUS mapSDocument() {

		URI resourceURI = getResourceURI();
		String resource = resourceURI.path();
		readSpreadsheetResource(resource);

		return (DOCUMENT_STATUS.COMPLETED);
	}

//	private StringBuilder currentText = new StringBuilder();
	
	STextualDS primaryText = null;
	// save all tokens of the current primary text
	List<SToken> currentTokList = new ArrayList<SToken>();

	private void readSpreadsheetResource(String resource) {

		// String primaryTextLayer = getProps().getPrimaryText();
		String primaryTextLayer = "tok";
		List<String> primaryTextLayerList = Arrays.asList(primaryTextLayer
				.split("\\s*,\\s*"));

		getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());

		SpreadsheetImporter.logger.debug("Importing the file {}.", resource);

		SpreadsheetImporter.logger.info(resource);
		InputStream xlsxFileStream;
		File xlsxFile;
		Workbook workbook = null;
		primaryText = getDocument().getDocumentGraph().createTextualDS("");

		try {
			xlsxFile = new File(resource);
			xlsxFileStream = new FileInputStream(xlsxFile);
			workbook = WorkbookFactory.create(xlsxFileStream);
			workbook.close();
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException e) {
			SpreadsheetImporter.logger.warn("Could not open file '" + resource
					+ "'.");
		}

		if (workbook != null) {

			int sheetNum = workbook.getNumberOfSheets();

			for (int currSheet = 0; currSheet < sheetNum; currSheet++) {
				Sheet sheet = workbook.getSheetAt(currSheet);
				// get the sheet holding the actual corpus data
				if (sheet.getSheetName().equals(getProps().getCorpusSheet())) {
					int rowNum = sheet.getLastRowNum();

					// get all names of the annotation layer (first row)
					Row headerRow = sheet.getRow(0);
					int columnNum = 0;
					if (headerRow != null) {
						columnNum = headerRow.getLastCellNum();
					}

					// iterate through all layers and save all layers
					// that hold the primary data
					List<Integer> primTextPos = new ArrayList<Integer>();
					for (int currCol = 0; currCol < columnNum; currCol++) {
						if (headerRow.getCell(currCol) != null) {
							if (primaryTextLayerList.contains(headerRow
									.getCell(currCol).toString())) {
								primTextPos.add(currCol);
							}
						}
					}

					String text = "";
					if (primTextPos != null) {

						int offset = primaryText.getText().length();

						for (int primText : primTextPos) {
							for (int currRow = 1; currRow <= rowNum; currRow++) {
								Row row = sheet.getRow(currRow);

								if (row != null
										&& row.getCell(primText) != null) {
									Cell primCell = row.getCell(primText);

									text = primCell.toString();
									int start = offset;
									int end = start + text.length();
									offset += text.length();

									getDocument().getDocumentGraph()
											.createToken(primaryText, start,
													end);

									// insert space between tokens
									if (currRow != rowNum
											|| primTextPos.indexOf(primText) != primTextPos
													.get(primTextPos.size() - 1)) {
										text += " ";
										offset++;
									}

									// TODO: delete debug print
									System.out.println(primCell.toString());

									primaryText.setText(primaryText.getText()
											+ text);
								}
								getDocument().getDocumentGraph().addNode(
										primaryText);
							}

						}
					}

				}

			}

		}

	}

}