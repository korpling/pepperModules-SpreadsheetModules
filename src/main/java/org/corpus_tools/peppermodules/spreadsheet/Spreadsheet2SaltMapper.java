package org.corpus_tools.peppermodules.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.pepper.modules.exceptions.PepperModuleException;
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.SOrderRelation;
import org.corpus_tools.salt.common.SSpan;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SLayer;
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

		return DOCUMENT_STATUS.COMPLETED;
	}
	// row that holds the names of each tier
	public Row headerRow;

	// save all tokens of the current primary text
	List<SToken> currentTokList = new ArrayList<>();

	SSpan tokSpan = SaltFactory.createSSpan();
	SSpan annoSpan = SaltFactory.createSSpan();
	
	// maybe use the StringBuilder
	private StringBuilder currentText = new StringBuilder();

	private Map<String, SLayer> sLayerMap = null;
	STimeline timeline = SaltFactory.createSTimeline();
	
	private void readSpreadsheetResource(String resource) {
		getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());
		
		getDocument().getDocumentGraph().setTimeline(timeline);
		
		SpreadsheetImporter.logger.debug("Importing the file {}.", resource);
		SpreadsheetImporter.logger.info(resource);

		// get the excel files here
		InputStream excelFileStream;
		File excelFile;
		Workbook workbook = null;

		try {
			excelFile = new File(resource);
			excelFileStream = new FileInputStream(excelFile);
			workbook = WorkbookFactory.create(excelFileStream);
			workbook.close();
			
			// put this into the code below
			Sheet corpusSheet;
			if (getProps().getCorpusSheet().equals("Tabelle1")) {
				corpusSheet = workbook.getSheetAt(0);
			} else {
				corpusSheet = workbook.getSheet(getProps().getCorpusSheet());
			}
			int lastRow = corpusSheet.getLastRowNum();
			mapCommonTimeLine2STimeline(lastRow, timeline);
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException e) {
			SpreadsheetImporter.logger.warn("Could not open file '" + resource
					+ "'.");
		}

		getPrimTextTiers(workbook);

		// setPrimText(column, primText, corpusSheet);
	}

	private void mapCommonTimeLine2STimeline(int lastRow, STimeline timeline) {
		int currRow = 1;
		while(currRow <= lastRow){
			timeline.increasePointOfTime();
			currRow++;
		}
	}

	private void getPrimTextTiers(Workbook workbook) {
		// get all primary layer
		String primaryTextLayer = getProps().getPrimaryText();
		// seperate string of primary text layer into list by commas
		List<String> primaryTextLayerList = Arrays.asList(primaryTextLayer
				.split("\\s*,\\s*"));

		if (workbook != null) {
			// get corpus sheet
			Sheet corpusSheet;
			if (getProps().getCorpusSheet().equals("Tabelle1")) {
				corpusSheet = workbook.getSheetAt(0);
			} else {
				corpusSheet = workbook.getSheet(getProps().getCorpusSheet());
			}

			if (corpusSheet != null) {

				// get all names of the annotation tiers (first row)
				headerRow = corpusSheet.getRow(0);
				// List for primary text and its annotations
				HashMap<Integer, List<Integer>> annoPrimRelations = new HashMap<>();
				// iterate through all tiers and save all tiers
				// that hold the primary data
				List<Integer> primTextPos = new ArrayList<Integer>();
				if (headerRow != null) {

					int currColumn = 0;
					while (currColumn < headerRow.getLastCellNum()
							&& headerRow.getCell(currColumn) != null
							&& !headerRow.getCell(currColumn).toString()
									.equals("")) {
						String tierName = headerRow.getCell(currColumn)
								.toString();
						if (primaryTextLayerList.contains(tierName)) {
							// save all indexes of layer containing primary text
							primTextPos.add(currColumn);
						} else {
							if (getPrimOfAnnoPrimRel(tierName) != null) {
								// current tier is an annotation and the
								// belonging primary text was set by property
								getAnnotatedPrim(
										getPrimOfAnnoPrimRel(tierName),
										annoPrimRelations, currColumn);
							} else if (tierName.matches(".+\\[.+\\]")) {
								// the belonging primary text was set by the
								// annotator
								String primTier = tierName.split("\\[")[1]
										.replace("]", "");
								getAnnotatedPrim(primTier, annoPrimRelations,
										currColumn);
							} else {
								SpreadsheetImporter.logger
										.warn("No primary text for the annotation '"
												+ tierName + "' given.");
							}
						}
						currColumn++;
					}
				}

				if (primTextPos != null) {

					// List<CellRangeAddress> allMergedRegions =
					// corpusSheet.getMergedRegions();
					// for(CellRangeAddress mergedRegion : allMergedRegions){
					// for (int primText : primTextPos) {
					// if(mergedRegion.getFirstColumn() != primText){
					//
					// }
					// }
					// }
					setPrimText(corpusSheet, primTextPos, annoPrimRelations);

				}

			}

		}
	}

	private void getAnnotatedPrim(String primTier,
			HashMap<Integer, List<Integer>> annoPrimRelations, int currColumn) {
		// get all annotations and their belonging primary text and save them
		// into a hashmap
		if (getColumn(primTier, headerRow) != null) {
			int primData = getColumn(primTier, headerRow);
			if (annoPrimRelations.get(primData) != null) {
				annoPrimRelations.get(primData).add(0, currColumn);
			} else {
				List<Integer> init = new ArrayList<Integer>();
				init.add(currColumn);
				annoPrimRelations.put(primData, init);
			}
		}
	}

	// computes, weather a cell of a given sheet is part of merged cells
	private Boolean isMergedCell(Cell currentCell, Sheet currentSheet) {
		Boolean isMergedCell = false;
		if(currentCell != null && currentSheet!= null){
		if (currentSheet.getNumMergedRegions() > 0) {
			List<CellRangeAddress> sumMergedRegions = currentSheet
					.getMergedRegions();
			for (CellRangeAddress mergedRegion : sumMergedRegions) {
				if (mergedRegion.isInRange(currentCell.getRowIndex(),
						currentCell.getColumnIndex())) {
					isMergedCell = true;
					break;
				}
			}
		}}
		return isMergedCell;
	}

	// returns the last cell of merged cells
	private int getLastCell(Cell primCell, Sheet currentSheet) {
		int lastCell = primCell.getRowIndex();
		List<CellRangeAddress> allMergedCells = currentSheet.getMergedRegions();
		for (CellRangeAddress mergedCell : allMergedCells) {
			if (mergedCell.isInRange(primCell.getRowIndex(),
					primCell.getColumnIndex())) {
				lastCell = mergedCell.getLastRow();
			}
		}
		return lastCell;
	}

	/**
	 * Set the primary text of each document
	 * 
	 * @param column
	 * @param primText
	 * @param corpusSheet
	 * @param primTextPos
	 * @param annoPrimRelations
	 * @return
	 */
	private void setPrimText(Sheet corpusSheet, List<Integer> primTextPos,
			HashMap<Integer, List<Integer>> annoPrimRelations) {
		STextualDS primaryText = null;
		// save all tokens of the current primary text
		for (int primText : primTextPos) {
			// initialize primaryText
			primaryText = SaltFactory.createSTextualDS();
			primaryText.setText("");
			getDocument().getDocumentGraph().addNode(primaryText);

			int offset = primaryText.getText().length();

			currentText = new StringBuilder();
			String text = currentText.toString();
			// (re-)initialize current tok list and span
			currentTokList = new ArrayList<>();
			tokSpan = SaltFactory.createSSpan();
			SToken lastTok = null;
			
			int currRow = 1;
			// the first condition may be redundant
			while (currRow < corpusSheet.getPhysicalNumberOfRows()
					&& currRow <= corpusSheet.getLastRowNum()) {
				Row row = corpusSheet.getRow(currRow);

				Cell primCell = row.getCell(primText);
				SToken currTok = null;

				if (primCell != null && !primCell.toString().equals("")) {
					text = primCell.toString();
					int start = offset;
					int end = start + text.length();
					offset += text.length();
					currentText.append(text);
					if (headerRow.getCell(primText) != null) {
						primaryText.setName(headerRow.getCell(primText)
								.toString());
					}

					currTok = getDocument().getDocumentGraph().createToken(
							primaryText, start, end);
				}

				if (lastTok != null && currTok != null) {
					SOrderRelation primTextOrder = SaltFactory
							.createSOrderRelation();
					primTextOrder.setSource(lastTok);
					primTextOrder.setTarget(currTok);
					getDocument().getDocumentGraph().addRelation(primTextOrder);
				}
				if (currTok != null) {
					// set timeline
					STimelineRelation sTimeRel = SaltFactory
							.createSTimelineRelation();

					sTimeRel.setSource(currTok);
					sTimeRel.setTarget(getDocument().getDocumentGraph()
							.getTimeline());
					int endTime = currRow;
					if (isMergedCell(primCell, corpusSheet)) {
						endTime = getLastCell(primCell, corpusSheet);
					}
					sTimeRel.setStart(currRow-1);
					// this is not correct?!
					sTimeRel.setEnd(endTime);
					getDocument().getDocumentGraph().addRelation(sTimeRel);

					// remember all SToken
					currentTokList.add(currTok);

					// insert space between tokens
					if (!primCell.toString().isEmpty()
							&& (currRow != corpusSheet.getLastRowNum())) {
						text += " ";
						offset++;
					}

					// TODO: delete debug print
					// System.out.println(text);
					primaryText.setText(primaryText.getText() + text);
				}
				if (currTok != null) {
					lastTok = currTok;
				}

				// insert this into the code above with a mapping on all the
				// annotation layers, to create the relations between tokens and
				// annotations
				//TODO: put this into a function and make sure, that there is a token list for every annotation tier but that is not recreated at every new row
				if (annoPrimRelations.get(primText) != null
						&& currentTokList != null) {
					List<Integer> annosOfCurPrim = annoPrimRelations
							.get(primText);
					for (int anno : annosOfCurPrim) {
						
						List<SToken> primForAnno = new ArrayList<>();
						if(currTok != null){
							primForAnno.add(currTok);
						}
						String annoName = headerRow.getCell(anno).toString();
						
						 if(isMergedCell(row.getCell(anno), corpusSheet)){
							 if(getLastCell(row.getCell(anno), corpusSheet) == currRow && primForAnno != null){
								 annoSpan = getDocument().getDocumentGraph().createSpan(primForAnno);
							 }
						 // store all token in a list
						// List<CellRangeAddress> sumMergedRegions =
						// corpusSheet.getMergedRegions();
						// for (CellRangeAddress mergedRegion :
						// sumMergedRegions){
						// if(mergedRegion.isInRange(row.getCell(anno).getRowIndex(),
						// row.getCell(anno).getColumnIndex())){
						// annotation =
						// corpusSheet.getRow(mergedRegion.getFirstRow()).getCell(anno).toString();
						// }
//						 }
						// if(getLastCell(row.getCell(anno), corpusSheet) ==
						// currRow){
						// tokSpan =
						// getDocument().getDocumentGraph().createSpan(currentTokList);
						// tokSpan.addAnnotation(getDocument().getDocumentGraph().createAnnotation(null,
						// annoName, annotation));
						 }
						// } else{
						// tokSpan =
						// getDocument().getDocumentGraph().createSpan(primForAnno);
						// annotation = row.getCell(anno).toString();
						// primForAnno = new ArrayList<>();
						// tokSpan.addAnnotation(getDocument().getDocumentGraph().createAnnotation(null,
						// annoName, annotation));
						// }
						//
						// }
					}
				}
				currRow++;
			}
				tokSpan = getDocument().getDocumentGraph().createSpan(
						currentTokList);
				tokSpan.setName(headerRow.getCell(primText).toString());
				getDocument().getDocumentGraph().addNode(tokSpan);
			}
			// for(int rowIndex = 1; rowIndex <= corpusSheet.getLastRowNum();
			// rowIndex++){
			// Row row = corpusSheet.getRow(rowIndex);
			// if (row != null && row.getCell(column) != null) {
			// Cell primCell = row.getCell(column);
			// primText.append(primCell.toString() + " ");
			// }
			// }
	
	}

	/**
	 * rename the annotation layer, so that the primary text, that the
	 * annotation layer refers to, is written behind the actual layer name
	 */
	private String getPrimOfAnnoPrimRel(String currentLayer) {
		String annoPrimRel = getProps().getAnnoPrimRel();
		String annoPrimNew = null;
		if (annoPrimRel != null) {
			List<String> annoPrimRelation = Arrays.asList(annoPrimRel
					.split("\\s*,\\s*"));
			for (String annoPrim : annoPrimRelation) {
				String annoName = annoPrim.split(">")[0];
				String annoPrimCouple = annoPrim.split(">")[1];

				if (annoName.equals(currentLayer)) {
					annoPrimNew = annoPrimCouple.split("\\[")[1].replace("]",
							"");
				}
			}
		}
		return annoPrimNew;
	}

	/**
	 * method to find a certain column, given by the row and the string of a
	 * cell
	 * 
	 * @param layerName
	 * @param row
	 * @return the column of a row, that includes the given string
	 */
	private Integer getColumn(String layerName, Row row) {
		Integer searchedColumn = null;
		for (int currCol = 0; currCol < row.getLastCellNum(); currCol++) {

			if (row.getCell(currCol) != null
					&& row.getCell(currCol).toString().equals(layerName)) {
				searchedColumn = currCol;
			}
		}
		return searchedColumn;
	}

	private Integer getRowOfLastToken(Sheet sheet, int coll) {
		int lastRow = sheet.getLastRowNum();
		for (int currRow = 1; currRow <= lastRow; currRow++) {
			Row row = sheet.getRow(currRow);
			if (row == null) {
				lastRow = currRow - 1;
				break;
			}
			if (row.getCell(coll) == null
					|| (row.getCell(coll).toString().equals("") && isMergedCell(
							row.getCell(coll), sheet))) {
				lastRow = currRow - 1;
				break;
			}
		}
		return lastRow;
	}

	private Map<String, SLayer> getLayerAnnoCouples() {
		// get all sLayer and the associated annotations
		// TODO: check for right syntax?
		String annoLayerCouple = getProps().getLayer();
		Map<String, SLayer> annoLayerCoupleMap = new HashMap<>();
		if (annoLayerCouple != null && !annoLayerCouple.isEmpty()) {
			List<String> annoLayerCoupleList = Arrays.asList(annoLayerCouple
					.split("\\s*},\\s*"));

			for (String annoLayer : annoLayerCoupleList) {
				List<String> annoLayerPair = Arrays
						.asList(annoLayer.split(">"));

				SLayer sLayer = SaltFactory.createSLayer();
				sLayer.setName(annoLayerPair.get(0));
				String[] assoAnno = annoLayerPair.get(1).split("\\s*,\\s*");
				for (String anno : assoAnno) {
					anno = anno.replace("{", "");
					anno = anno.replace("}", "");
					annoLayerCoupleMap.put(anno, sLayer);

					// System.out.println(sLayer.getName() + ": " + anno);
				}
			}
		}
		else {
			throw new PepperModuleException(
				"Cannot import the given data, because property file contains a corrupt value for property '"
					+ getProps().getLayer()
					+ "'. Please check the brackets/ syntax you used.");
			}
		
		return annoLayerCoupleMap;
	}
	
//	private int getLastRowOfCorpus(Workbook workbook) {
//		return 0;
//		
//	}

	// private void setAnnotations(SSpan tokSpan){
	// for (int currRow = 1; currRow <= rowNum; currRow++) {
	// Row row = sheet.getRow(currRow);
	//
	// if (row != null
	// && row.getCell(primText) != null) {
	// Cell primCell = row.getCell(primText);
	// }
	// }
	// }

	// TODO: add function to check for every token, if the associated
	// annotations are merged cells -> if not: add annotation to token,
	// otherwise find the last cell of the merged cells, on token layer, check
	// if line = line of last cell -> create a span of all token between start
	// and end point and add annotation to the span.

}