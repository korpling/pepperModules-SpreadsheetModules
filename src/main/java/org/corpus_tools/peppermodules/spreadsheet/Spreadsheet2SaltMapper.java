/**
 * Copyright 2009 Humboldt-Universität zu Berlin, INRIA.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 *
 */
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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.salt.SALT_TYPE;
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.SOrderRelation;
import org.corpus_tools.salt.common.SSpan;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SLayer;
import org.corpus_tools.salt.core.SNode;
import org.corpus_tools.salt.util.DataSourceSequence;
import org.eclipse.emf.common.util.URI;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import com.google.common.collect.Tables;

/**
 * 
 * @author Vivian Voigt
 *
 */
public class Spreadsheet2SaltMapper extends PepperMapperImpl implements PepperMapper {

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


	private void readSpreadsheetResource(String resource) {
		getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());

		STimeline timeline = SaltFactory.createSTimeline();
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

		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			SpreadsheetImporter.logger.warn("Could not open file '" + resource + "'.");
		}

		getPrimTextTiers(workbook, timeline);

	}

	/**
	 * Map the line number of a given token to the {@link STimeline}
	 * 
	 * @param lastRow
	 * @param timeline
	 */
	private void mapLinenumber2STimeline(int lastRow, STimeline timeline) {
		int currRow = 1;
		while (currRow <= lastRow) {
			timeline.increasePointOfTime();
			currRow++;
		}
	}
	
	/**
	 * Create a read-only lookup table for the merged cells.
	 * @param mergedCells
	 * @return
	 */
	private Table<Integer, Integer, CellRangeAddress> calculateMergedCellIndex(
			List<CellRangeAddress> mergedCells) {
		
		Table<Integer, Integer, CellRangeAddress> idx = HashBasedTable.create();		
		if(mergedCells != null) {
			for(CellRangeAddress cell : mergedCells) {
				// add each underlying row/column of the range
				for(int i=cell.getFirstRow(); i <= cell.getLastRow(); i++) {
					for(int j=cell.getFirstColumn(); j <= cell.getLastColumn(); j++) {
						idx.put(i, j, cell);
					}
				}
			}
		}		
		return Tables.unmodifiableTable(idx);
	}

	/**
	 * get the primary text tiers of the given document
	 * 
	 * @param workbook
	 *            of the excel file
	 */
	private void getPrimTextTiers(Workbook workbook, STimeline timeline) {
		// get all primary text tiers
		String primaryTextTier = getProps().getPrimaryText();
		// seperate string of primary text tiers into list by commas
		List<String> primaryTextTierList = Arrays.asList(primaryTextTier.split("\\s*,\\s*"));

		if (workbook != null) {
			// get corpus sheet
			Sheet corpusSheet;
			// default ("Tabelle1"/ first sheet)
			if (getProps().getCorpusSheet().equals("Tabelle1")) {
				corpusSheet = workbook.getSheetAt(0);
			} else {
				// get corpus sheet by name
				corpusSheet = workbook.getSheet(getProps().getCorpusSheet());
			}
			// end of the excel file
			int lastRow = corpusSheet.getLastRowNum();
			mapLinenumber2STimeline(lastRow, timeline);

			if (corpusSheet != null) {

				// row with all names of the annotation tiers (first row)
				Row headerRow = corpusSheet.getRow(0);
				// List for each primary text and its annotations
				HashMap<Integer, List<Integer>> annoPrimRelations = new HashMap<>();

				List<Integer> primTextPos = new ArrayList<Integer>();
				if (headerRow != null) {
					// iterate through all tiers and save tiers (column number)
					// that hold the primary data

					int currColumn = 0;
					while (currColumn < headerRow.getLastCellNum() && headerRow.getCell(currColumn) != null
							&& !headerRow.getCell(currColumn).toString().equals("")) {
						String tierName = headerRow.getCell(currColumn).toString();
						if (primaryTextTierList.contains(tierName)) {
							// current tier contains primary text
							// save all indexes of tier containing primary text
							primTextPos.add(currColumn);
						} else {
							// current tier contains (other) annotations
							if (getPrimOfAnnoPrimRel(tierName) != null) {
								// current tier is an annotation and the
								// belonging primary text was set by property
								setAnnotationPrimCouple(getPrimOfAnnoPrimRel(tierName), annoPrimRelations, currColumn, headerRow);
							} else if (tierName.matches(".+\\[.+\\]")) {
								// the belonging primary text was set by the
								// annotator
								String primTier = tierName.split("\\[")[1].replace("]", "");
								setAnnotationPrimCouple(primTier, annoPrimRelations, currColumn, headerRow);
								
							} else if (primaryTextTierList.size() == 1 && getProps().getAnnoPrimRel() == null) {
								// There is only one primary text so we can safely assume this is the one
								// the annotation is connected to.
								setAnnotationPrimCouple(primaryTextTierList.get(0), annoPrimRelations, currColumn, headerRow);
							} else {
								SpreadsheetImporter.logger
										.warn("No primary text for the annotation '" + tierName + "' given.");
							}
						}
						currColumn++;
					}
				}

				if (primTextPos != null) {

					setPrimText(corpusSheet, primTextPos, annoPrimRelations, headerRow);
				}
			}
			if (getProps().getMetaAnnotation()) {
				setDocMetaData(workbook);
			}
		}
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
			HashMap<Integer, List<Integer>> annoPrimRelations, Row headerRow) {
		// initialize with number of token we have to create
		int progressTotalNumberOfColumns = primTextPos.size();
		// add each annotation to this number
		for(List<Integer> annos : annoPrimRelations.values()) {
			progressTotalNumberOfColumns += annos.size();
		}
		int progressProcessedNumberOfColumns = 0;
		
		final Map<String, SLayer> layerTierCouples = getLayerTierCouples();
		final Table<Integer, Integer, CellRangeAddress> mergedCells = 
				calculateMergedCellIndex(corpusSheet.getMergedRegions());
		
		// use formater to ensure that e.g. integers will not be converted into decimals
		DataFormatter formatter = new DataFormatter();
		// save all tokens of the current primary text
		List<SToken> currentTokList = new ArrayList<>();
		// save all tokens of the current primary text
		for (int primText : primTextPos) {

			// initialize primaryText
			STextualDS primaryText = SaltFactory.createSTextualDS();
			StringBuilder currentText = new StringBuilder();
			
			
			if (headerRow.getCell(primText) != null) {
				primaryText.setName(headerRow.getCell(primText).toString());
			}
			getDocument().getDocumentGraph().addNode(primaryText);

			int offset = currentText.length();

			
			SToken lastTok = null;

			// start with the second row of the table, since the first row holds
			// the name of each tier
			int currRow = 1;
			while (currRow < corpusSheet.getPhysicalNumberOfRows()) {
				// iterate through all rows of the given corpus sheet

				Row row = corpusSheet.getRow(currRow);
				Cell primCell = row.getCell(primText);
				SToken currTok = null;
				int endCell = currRow;

				String text = null;
				if (primCell != null && !primCell.toString().isEmpty()) {
					text = formatter.formatCellValue(primCell);
					
				} else if(getProps().getIncludeEmptyPrimCells()) {
					text = "";
								
				}
				if(text != null){
					int start = offset;
					int end = start + text.length();
					offset += text.length();
					currentText.append(text);

					currTok = getDocument().getDocumentGraph().createToken(primaryText, start, end);
					
					if(primCell != null){
						endCell = getLastCell(primCell, mergedCells);
					}
				}
				
				
				if(currTok != null) {
					if (lastTok != null && getProps().getAddOrderRelation()) {
						addOrderRelation(lastTok, currTok, headerRow.getCell(primText).toString());
					}
					// add timeline relation
					addTimelineRelation(currTok, currRow, endCell, corpusSheet);

					// remember all SToken
					currentTokList.add(currTok);
					
					// insert space between tokens
					if (text != null && (currRow != corpusSheet.getLastRowNum())) {
						currentText.append(" ");
						offset++;
					}
				}

				
				if (currTok != null) {
					lastTok = currTok;
				}
				currRow++;
			} // end for each token row
			primaryText.setText(currentText.toString());
			
			progressProcessedNumberOfColumns++;
			setProgress((double) progressProcessedNumberOfColumns / (double) progressTotalNumberOfColumns);

			if (getProps().getLayer() != null) {

				if (currentTokList != null && layerTierCouples.size() > 0) {
					if (layerTierCouples.get(primaryText.getName()) != null) {
						SLayer sLayer = layerTierCouples.get(primaryText.getName());
						getDocument().getDocumentGraph().addLayer(sLayer);
						for(SToken t : currentTokList) {
							sLayer.addNode(t);
						}
					}
				}
			}

			if (annoPrimRelations.get(primText) != null) {
				for (int annoTier : annoPrimRelations.get(primText)) {

					SSpan annoSpan = null;
					int currAnno = 1;
					
					while (currAnno < corpusSheet.getPhysicalNumberOfRows()) {
						Row row = corpusSheet.getRow(currAnno);
						Cell annoCell = row.getCell(annoTier);

						if (annoCell != null && !annoCell.toString().equals("")) {
							String annoText = "";
							annoText = formatter.formatCellValue(annoCell);
							int annoStart = currAnno - 1;
							int annoEnd = getLastCell(annoCell, mergedCells);
							DataSourceSequence<Integer> sequence = new DataSourceSequence<Integer>();
							sequence.setStart(annoStart);
							sequence.setEnd(annoEnd);
							sequence.setDataSource(getDocument().getDocumentGraph().getTimeline());

							List<SToken> sTokens = getDocument().getDocumentGraph().getTokensBySequence(sequence);

							List<SToken> tokenOfSpan = new ArrayList<>();

							if(sTokens == null) {
								SpreadsheetImporter.logger.error("Segmentation error: The segmentation of the tier \""
										+ headerRow.getCell(annoTier).toString() + "\" in the document: \""
										+ getResourceURI().lastSegment() + "\" in line: " + currAnno
										+ " does not match to its primary text: \""
										+ headerRow.getCell(primText).toString() + "\".");
							} else {
								for (SToken tok : sTokens) {
									STextualDS textualDS = getTextualDSForNode(tok, getDocument().getDocumentGraph());
									if (textualDS.getName().equals(headerRow.getCell(primText).toString())) {
										tokenOfSpan.add(tok);
									}
								}
							}

							annoSpan = getDocument().getDocumentGraph().createSpan(tokenOfSpan);

							if (annoSpan != null && headerRow.getCell(annoTier) != null
									&& !headerRow.getCell(annoTier).toString().equals("")) {
								annoSpan.createAnnotation(null, headerRow.getCell(annoTier).toString(), annoText);
								annoSpan.setName(headerRow.getCell(annoTier).toString());
							}

						}

						currAnno++;
					} // end for each row of annotation
					if (getProps().getLayer() != null && annoSpan != null) {

						if (layerTierCouples.size() > 0) {
							if (layerTierCouples.get(headerRow.getCell(annoTier).toString()) != null) {
								SLayer sLayer = layerTierCouples.get(headerRow.getCell(annoTier).toString());
								getDocument().getDocumentGraph().addLayer(sLayer);
								sLayer.addNode(annoSpan);
							}
						}
					}
					progressProcessedNumberOfColumns++;
					setProgress((double) progressProcessedNumberOfColumns / (double) progressTotalNumberOfColumns);
				} // end for each annotation layer
			}
		} // end for each primTextPos
	}
	
	private void addTimelineRelation(SToken tok, int currRow, int endTime, Sheet corpusSheet) {
		STimelineRelation sTimeRel = SaltFactory.createSTimelineRelation();
		sTimeRel.setSource(tok);
		sTimeRel.setTarget(getDocument().getDocumentGraph().getTimeline());
		sTimeRel.setStart(currRow - 1);
		sTimeRel.setEnd(endTime);
		getDocument().getDocumentGraph().addRelation(sTimeRel);
	}
	
	private void addOrderRelation(SToken lastTok, SToken currTok, String name) {
		SOrderRelation primTextOrder = SaltFactory.createSOrderRelation();
		primTextOrder.setType(name);

		primTextOrder.setSource(lastTok);
		primTextOrder.setTarget(currTok);
		getDocument().getDocumentGraph().addRelation(primTextOrder);
	}

	/**
	 * get all annotations and their belonging primary text and save them into a
	 * hashmap
	 * 
	 * @param primTier
	 * @param annoPrimRelations
	 * @param currColumn
	 */
	private void setAnnotationPrimCouple(String primTier, HashMap<Integer, List<Integer>> annoPrimRelations,
			int currColumn, Row headerRow) {

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

	/**
	 * Return the last cell of merged cells
	 * 
	 * @param primCell,
	 *            current cell of the primary text
	 * @return
	 */
	private int getLastCell(Cell cell, Table<Integer, Integer, CellRangeAddress> mergedCellsIdx) {
		int lastCell = cell.getRowIndex();
		CellRangeAddress mergedCell = mergedCellsIdx.get(cell.getRowIndex(), cell.getColumnIndex());
		if(mergedCell != null) {
			lastCell = mergedCell.getLastRow();
		}
		return lastCell;
	}

	/**
	 * rename the annotation tier, so that the primary text, that the annotation
	 * tier refers to, is written behind the actual tier name
	 * 
	 * @param currentTier
	 */
	private String getPrimOfAnnoPrimRel(String currentTier) {
		String annoPrimNew = null;
		
		// check if both properties where used
		if(getProps().getAnnoPrimRel() != null && getProps().getShortAnnoPrimRel() != null){
			SpreadsheetImporter.logger.error("Wrong property handling. Please use only one property to specify which annotation refers to which primary text tier (exclusive use of either 'annoPrimRel' or 'shortAnnoPrimRel').");
		} else{
		String annoPrimRel = getProps().getAnnoPrimRel();
		if (annoPrimRel != null) {
			List<String> annoPrimRelation = Arrays.asList(annoPrimRel.split("\\s*,\\s*"));
			for (String annoPrim : annoPrimRelation) {
				String[] splitted = annoPrim.split("=", 2);
				if(splitted.length > 1 ) {
					String annoName = splitted[0];
					String annoPrimCouple = splitted[1];
	
					if (annoName.equals(currentTier)) {
						annoPrimNew = annoPrimCouple.split("\\[")[1].replace("]", "");
					}
				}
			}
		}
		String shortAnnoPrimRel = getProps().getShortAnnoPrimRel();
		if(shortAnnoPrimRel != null){
		
			List<String> shortAnnoPrimRelation = Arrays.asList(shortAnnoPrimRel.split("\\s*},\\s*"));
			for (String annoPrim : shortAnnoPrimRelation) {
				List<String> primAnnoPair = Arrays.asList(annoPrim.split("="));
				String[] annos = primAnnoPair.get(1).split("\\s*,\\s*");
				
					for(String anno : annos){
						if (anno.replace("}", "").replace("{", "").equals(currentTier)){
							annoPrimNew = primAnnoPair.get(0);
						}
					}
					System.out.println(currentTier +", " +annoPrimNew);
				}
			}
		}
		return annoPrimNew;
	}

	/**
	 * method to find a certain column, given by the row and the string of a
	 * cell
	 * 
	 * @param tierName
	 * @param row
	 * @return the column of a row, that includes the given string
	 */
	private Integer getColumn(String tierName, Row row) {
		Integer searchedColumn = null;
		for (int currCol = 0; currCol < row.getLastCellNum(); currCol++) {

			if (row.getCell(currCol) != null && row.getCell(currCol).toString().equals(tierName)) {
				searchedColumn = currCol;
			}
		}
		return searchedColumn;
	}

	/**
	 * Finds the {@link STextualDS} for a given node. The node must dominate a
	 * token of this text. (copied from
	 * https://github.com/korpling/ANNIS/blob/develop/annis-interfaces/src/main/
	 * java/annis/CommonHelper.java#L359)
	 *
	 * @param node
	 * @return
	 */
	public static STextualDS getTextualDSForNode(SNode node, SDocumentGraph graph) {
		if (node != null) {
			List<DataSourceSequence> dataSources = graph.getOverlappedDataSourceSequence(node,
					SALT_TYPE.STEXT_OVERLAPPING_RELATION);
			if (dataSources != null) {
				for (DataSourceSequence seq : dataSources) {
					if (seq.getDataSource() instanceof STextualDS) {
						return (STextualDS) seq.getDataSource();
					}
				}
			}
		}
		return null;
	}

	private Map<String, SLayer> getLayerTierCouples() {
		// get all sLayer and the associated annotations
		String tierLayerCouple = getProps().getLayer();
		// System.out.println(annoLayerCouple);
		Map<String, SLayer> annoLayerCoupleMap = new HashMap<>();
		if (tierLayerCouple != null && !tierLayerCouple.isEmpty()) {
			List<String> annoLayerCoupleList = Arrays.asList(tierLayerCouple.split("\\s*},\\s*"));
			// System.out.println(annoLayerCoupleList);

			for (String annoLayer : annoLayerCoupleList) {
				List<String> annoLayerPair = Arrays.asList(annoLayer.split("="));

				SLayer sLayer = SaltFactory.createSLayer();
				sLayer.setName(annoLayerPair.get(0));
				String[] assoAnno = annoLayerPair.get(1).split("\\s*,\\s*");
				for (String tier : assoAnno) {
					tier = tier.replace("{", "");
					tier = tier.replace("}", "");
					annoLayerCoupleMap.put(tier, sLayer);
				}
			}
		}
		// TODO: throw an exception, if the syntax of the property is false
		// else {
		// throw new PepperModuleException(
		// "Cannot import the given data, because the property file contains a
		// corrupt value for property '"
		// + getProps().getLayer()
		// + "'. Please check the brackets/ syntax you used.");
		// }

		return annoLayerCoupleMap;
	}

	private void setDocMetaData(Workbook workbook) {
		Sheet metaSheet = null;

		// default ("Tabelle2"/ second sheet)
		if (getProps().getMetaSheet().equals("Tabelle2")) {
			if(workbook.getNumberOfSheets() > 1) {
				metaSheet = workbook.getSheetAt(1);
			}
		} else {
			// get corpus sheet by name
			metaSheet = workbook.getSheet(getProps().getCorpusSheet());
		}

		if (metaSheet != null) {
			DataFormatter formatter = new DataFormatter();
			// start with the second row of the table, since the first row holds
			// the name of each tier
			int currRow = 1;
			while (currRow < metaSheet.getPhysicalNumberOfRows()) {
				// iterate through all rows of the given meta informations

				Row row = metaSheet.getRow(currRow);
				Cell metaKey = row.getCell(0);
				Cell metaValue = row.getCell(1);

				if (metaKey != null && !metaKey.toString().equals("")) {
					if (metaValue != null && !metaValue.toString().equals("")) {
						if (getDocument().getMetaAnnotation(metaKey.toString()) == null) {
							getDocument().createMetaAnnotation(null, formatter.formatCellValue(metaKey),
									formatter.formatCellValue(metaValue));
						} else {
							SpreadsheetImporter.logger
									.warn("A meta information with the name \"" + formatter.formatCellValue(metaKey)
											+ "\" allready exists and will not be replaced.");
						}
					} else {
						SpreadsheetImporter.logger
								.warn("No value for the meta data: \"" + metaKey.toString() + "\" found.");
					}
				}
				currRow++;
			}
		}
	}

}
