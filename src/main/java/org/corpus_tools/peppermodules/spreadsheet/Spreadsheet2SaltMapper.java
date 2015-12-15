package org.corpus_tools.peppermodules.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

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
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.SOrderRelation;
import org.corpus_tools.salt.common.SSpan;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SAnnotation;
import org.corpus_tools.salt.core.SRelation;
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
	
	public Row headerRow;
	
	// save all tokens of the current primary text
	List<SToken> currentTokList = new ArrayList<>();

	SSpan tokSpan = SaltFactory.createSSpan();
	// maybe use the StringBuilder
	private StringBuilder currentText = new StringBuilder();
	private STextualDS primaryText = null;
	
	private void readSpreadsheetResource(String resource) {

		String primaryTextLayer = getProps().getPrimaryText();
		// seperate string of primary text layer into list
		List<String> primaryTextLayerList = Arrays.asList(primaryTextLayer
				.split("\\s*,\\s*"));

		SpreadsheetImporter.logger.debug("Importing the file {}.", resource);

		SpreadsheetImporter.logger.info(resource);
		
		// get the excel files here
		InputStream excelFileStream;
		File excelFile;
		Workbook workbook = null;
//		getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());
		
		try {
			excelFile = new File(resource);
			excelFileStream = new FileInputStream(excelFile);
			workbook = WorkbookFactory.create(excelFileStream);
			workbook.close();
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException e) {
			SpreadsheetImporter.logger.warn("Could not open file '" + resource
					+ "'.");
		}

		if (workbook != null) {

			int sheetNum = workbook.getNumberOfSheets();
			// iterate through sheets and find the one, that holds the corpus data
			for (int currSheet = 0; currSheet < sheetNum; currSheet++) {
				
				Sheet sheet = workbook.getSheetAt(currSheet);
				
				// get the sheet holding the actual corpus data
				if (sheet.getSheetName().equals(getProps().getCorpusSheet())) {
					currentText = new StringBuilder();
					
					int rowNum = sheet.getLastRowNum();

					// get all names of the annotation layer (first row)
					headerRow = sheet.getRow(0);
					int columnNum = 0;
					if (headerRow != null) {
						columnNum = headerRow.getLastCellNum();
					}

					HashMap<Integer, List<Integer>> annoPrimRelations = new HashMap<>();
					// iterate through all layers and save all layers
					// that hold the primary data
					List<Integer> primTextPos = new ArrayList<Integer>();
					for (int currColumn = 0; currColumn < columnNum; currColumn++) {
						
						if(headerRow.getCell(currColumn) == null || headerRow.getCell(currColumn).toString().equals("")){
							break;
						}
						
						if (headerRow.getCell(currColumn) != null) {
							String layerName = headerRow.getCell(currColumn).toString();
							if (primaryTextLayerList.contains(layerName)) {
								// save all indexes of layer containing primary text
								primTextPos.add(currColumn);
							}else{
								if(getPrimOfAnnoPrimRel(layerName) != null){
									// the belonging primary text was set by property
									if(getColumn(getPrimOfAnnoPrimRel(layerName), headerRow) != null){
										// get all annotations and their belonging primary text and save them into a hashmap
										int primData = getColumn(getPrimOfAnnoPrimRel(layerName), headerRow);
										if(annoPrimRelations.get(primData) != null){
											annoPrimRelations.get(primData).add(0, currColumn);
										} else{
											List<Integer> init = new ArrayList<Integer>();
											init.add(currColumn);
											annoPrimRelations.put(primData, init);
										}
									}
								} else if(layerName.matches(".+\\[.+\\]")){
									// the belonging primary text was set by the annotator
									String primLayer = layerName.split("\\[")[1].replace("]", "");
									// get all annotations and their belonging primary text and save them into a hashmap
									// TODO: write a method for this (it's reused)
									if(getColumn(primLayer, headerRow) != null){
										int primData = getColumn(primLayer, headerRow);
										if(annoPrimRelations.get(primData) != null){
											annoPrimRelations.get(primData).add(0, currColumn);
										} else{
											List<Integer> init = new ArrayList<Integer>();
											init.add(currColumn);
											annoPrimRelations.put(primData, init);
										}
									}
								} else{
									SpreadsheetImporter.logger.warn("No primary text for the annotation '" + layerName + "' given.");
								}
							}
						}
					}
					
					String text = "";
					if (primTextPos != null) {
						if (primaryText == null) {
							// initialize primaryText

							primaryText = SaltFactory.createSTextualDS();
							primaryText.setText("");
							getDocument().getDocumentGraph().addNode(primaryText);
						}
						int offset = primaryText.getText().length();

						STimeline timeline = SaltFactory.createSTimeline();
						
						getDocument().getDocumentGraph().setTimeline(timeline);				
						for (int primText : primTextPos) {
							// (re-) initialize current tok list and span
							currentTokList = new ArrayList<>();
							tokSpan = SaltFactory.createSSpan();
							SToken lastTok = null;
							
							for (int currRow = 1; currRow <= rowNum; currRow++) {
								
								Row row = sheet.getRow(currRow);
								
								if(row == null || row.getCell(primText) == null || (row.getCell(primText).toString().equals("") && !isMergedCell(row.getCell(primText), sheet))){
									break;
								}
								
								System.out.println("row: " + currRow + "/" + rowNum);
								
								if (row != null
										&& row.getCell(primText) != null) {
									Cell primCell = row.getCell(primText);

									SToken currTok = null;
									if (primCell != null && !primCell.toString().equals("")){
										text = primCell.toString();
										int start = offset;
										int end = start + text.length();
										offset += text.length();

										if(headerRow.getCell(primText) != null){
										primaryText.setName(headerRow.getCell(primText).toString());
										}
										
										currTok = getDocument().getDocumentGraph()
												.createToken(primaryText, start,
														end);
									}

									if (lastTok != null && currTok != null){
										SOrderRelation primTextOrder = SaltFactory.createSOrderRelation();
										primTextOrder.setSource(lastTok);
										primTextOrder.setTarget(currTok);
										getDocument().getDocumentGraph().addRelation(primTextOrder);
									}
									
									if(currTok != null){
										
										// set timeline
										STimelineRelation sTimeRel = SaltFactory.createSTimelineRelation();
										
										sTimeRel.setSource(currTok);
										sTimeRel.setTarget(getDocument().getDocumentGraph().getTimeline());
										int endTime = currRow;
										if(isMergedCell(primCell, sheet)){
											endTime = getLastCell(primText, sheet);
										}
										sTimeRel.setStart(currRow -1);
										System.out.println("tok: " + currTok.toString() + " text: " + text + ", startTime: " + (currRow -1)+", endTime: " + endTime);
										sTimeRel.setEnd(endTime);
										getDocument().getDocumentGraph().addRelation(sTimeRel);
										
										// remember all SToken
										currentTokList.add(currTok);
										
										// insert space between tokens
										if (!primCell.toString().isEmpty() && (currRow != rowNum
												|| primTextPos.indexOf(primText) != primTextPos.size() - 1)) {
											text += " ";
											offset++;
										}

										// TODO: delete debug print
										System.out.println(text);

										primaryText.setText(primaryText.getText()
												+ text);
									}
									if (currTok != null){
										lastTok = currTok;
									}	
							}
								
								// create a span for the current primary text
						}
							tokSpan = getDocument().getDocumentGraph().createSpan(currentTokList);
							
							// insert this into the code above with a mapping on all the annotation layers, to create the relations between tokens and annotations
							if(annoPrimRelations.get(primText) != null){
								List<Integer> annosOfCurPrim = annoPrimRelations.get(primText);
								for (int anno : annosOfCurPrim){
									
								}
							}
						}
					}
				}
			}

		}

	}
	
	// computes, weather a cell of a given sheet is part of merged cells
	private Boolean isMergedCell(Cell currentCell, Sheet currentSheet){
		Boolean isMergedCell = false;
		if (currentSheet.getNumMergedRegions() > 0){
			List<CellRangeAddress> sumMergedRegions = currentSheet.getMergedRegions();
			for (CellRangeAddress mergedRegion : sumMergedRegions){
				if(mergedRegion.isInRange(currentCell.getRowIndex(), currentCell.getColumnIndex())){
					isMergedCell = true;
					break;
				}
			}
		} 
		return isMergedCell;
	}
	
	// returns the last cell of a merged cell
	private int getLastCell(int currentCellIndex, Sheet currentSheet){
		return currentSheet.getMergedRegion(currentCellIndex).getLastRow();
	}
	
	
	private StringBuilder getPrimText(int column, StringBuilder primText, Sheet corpusSheet){
		for(int rowIndex = 1; rowIndex <= corpusSheet.getLastRowNum(); rowIndex++){
			Row row = corpusSheet.getRow(rowIndex);
			if (row != null && row.getCell(column) != null) {
				Cell primCell = row.getCell(column);
				primText.append(primCell.toString() + " ");					
			}
		}
		return null;
	}
	
	/**
	 * rename the annotation layer, so that the primary text, that the annotation layer refers to, is wrote behind the actual layer name
	 */
	private String getPrimOfAnnoPrimRel(String currentLayer){
		String annoPrimRel = getProps().getAnnoPrimRel();
		String annoPrimNew = null;
		if(annoPrimRel != null){
			List<String> annoPrimRelation = Arrays.asList(annoPrimRel
					.split("\\s*,\\s*"));
			for(String annoPrim : annoPrimRelation){
				String annoName = annoPrim.split(">")[0];
				String annoPrimCouple = annoPrim.split(">")[1];
				
				if(annoName.equals(currentLayer)){
					annoPrimNew = annoPrimCouple.split("\\[")[1].replace("]", "");
				}
			}
		}
		return annoPrimNew;
	}
	
	/**
	 * method to find a certain column, given by the row and the string of a cell
	 * @param layerName
	 * @param row
	 * @return the column of a row, that includes the given string
	 */
	private Integer getColumn(String layerName, Row row){
		Integer searchedColumn = null;
		for(int currCol = 0; currCol < row.getLastCellNum(); currCol++){
			
			if(row.getCell(currCol) != null && row.getCell(currCol).toString().equals(layerName)){
				searchedColumn = currCol;
			}
		}
		return searchedColumn;
	}
	
	private Integer getRowOfLastToken(Sheet sheet, int coll){
		int lastRow = sheet.getLastRowNum();
		for(int currRow = 1; currRow <= lastRow; currRow++){
			Row row = sheet.getRow(currRow);
			if(row == null){
				lastRow = currRow-1;
				break;
			}
			if(row.getCell(coll) == null || (row.getCell(coll).toString().equals("") && isMergedCell(row.getCell(coll), sheet))){
				lastRow = currRow -1;
				break;
			}
		}
		return lastRow;
	}
	
//	private void setAnnotations(SSpan tokSpan){
//		for (int currRow = 1; currRow <= rowNum; currRow++) {
//			Row row = sheet.getRow(currRow);
//
//			if (row != null
//					&& row.getCell(primText) != null) {
//				Cell primCell = row.getCell(primText);
//			}
//		}
//	}
	
}