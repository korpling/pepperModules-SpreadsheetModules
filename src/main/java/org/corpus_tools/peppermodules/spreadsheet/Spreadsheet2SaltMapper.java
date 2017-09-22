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

import com.google.common.base.Joiner;
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
import org.apache.poi.ss.util.CellReference;
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
import java.util.TreeSet;

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

	/**
	 * open document, create document graph and timeline in salt, print logging
	 * infos, throw warning, if there are any problems while handling the
	 * document
	 * 
	 * @param resource
	 *            string of the document path
	 */
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

		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException e) {
			SpreadsheetImporter.logger.warn("Could not open file '" + resource
					+ "'.");
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
	 * 
	 * @param mergedCells
	 * @return
	 */
	private Table<Integer, Integer, CellRangeAddress> calculateMergedCellIndex(
			List<CellRangeAddress> mergedCells) {

		Table<Integer, Integer, CellRangeAddress> idx = HashBasedTable.create();
		if (mergedCells != null) {
			for (CellRangeAddress cell : mergedCells) {
				// add each underlying row/column of the range
				for (int i = cell.getFirstRow(); i <= cell.getLastRow(); i++) {
					for (int j = cell.getFirstColumn(); j <= cell
							.getLastColumn(); j++) {
						idx.put(i, j, cell);
					}
				}
			}
		}
		return Tables.unmodifiableTable(idx);
	}

	/**
	 * get the primary text tiers and their annotations of the given document
	 * 
	 * @param workbook
	 * @param timeline
	 */
	private void getPrimTextTiers(Workbook workbook, STimeline timeline) {
		// get all primary text tiers
		String primaryTextTier = getProps().getPrimaryText();
		// seperate string of primary text tiers into list by commas
		List<String> primaryTextTierList = Arrays.asList(primaryTextTier
				.split("\\s*,\\s*"));
        
        TreeSet<String> annosWithoutPrim = new TreeSet<>();

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
				HashMap<Integer, Integer> annoPrimRelations = new HashMap<>();

				List<Integer> primTextPos = new ArrayList<Integer>();
				if (headerRow != null) {


                    // iterate through all tiers and save tiers (column number)
					// that hold the primary data

					int currColumn = 0;
                    

					List<String> emptyColumnList = new ArrayList<>();
					while (currColumn < headerRow.getPhysicalNumberOfCells()) {
						if (headerRow.getCell(currColumn) == null
								|| headerRow.getCell(currColumn).toString()
										.isEmpty()) {
							String emptyColumn = CellReference
									.convertNumToColString(currColumn);
							emptyColumnList.add(emptyColumn);
							currColumn++;
							continue;
						} else {
							if (!emptyColumnList.isEmpty()) {
								for (String emptyColumn : emptyColumnList) {
									SpreadsheetImporter.logger.warn("Column \""
											+ emptyColumn + "\" in document \""
											+ getResourceURI().lastSegment()
											+ "\" has no name.");
								}
								emptyColumnList = new ArrayList<>();
							}
                            
                            boolean primWasFound = false;
                            
							String tierName = headerRow.getCell(currColumn)
									.toString();
							if (primaryTextTierList.contains(tierName)) {
								// current tier contains primary text
								// save all indexes of tier containing primary
								// text
								primTextPos.add(currColumn);
                                primWasFound = true;
							} else {
								// current tier contains (other) annotations
								if (tierName.matches(".+\\[.+\\]")
										|| getProps().getAnnoPrimRel() != null
										|| getProps().getShortAnnoPrimRel() != null) {

									if (tierName.matches(".+\\[.+\\]")) {
										// the belonging primary text was set by
										// the annotator
										String primTier = tierName.split("\\[")[1]
												.replace("]", "");
										setAnnotationPrimCouple(primTier,
												annoPrimRelations, currColumn,
												headerRow);
                                        primWasFound = true;
									}
                                    
                                    String primOfAnnoFromConfig = 
                                            getPrimOfAnnoPrimRel(tierName
											.split("\\[")[0]);

									if (primOfAnnoFromConfig != null) {
										// current tier is an annotation and the
										// belonging primary text was set by
										// property
										setAnnotationPrimCouple(
												primOfAnnoFromConfig,
												annoPrimRelations, currColumn,
												headerRow);
                                        primWasFound = true;
									}
                                    
								} else if (primaryTextTierList.size() == 1
										&& getProps().getAnnoPrimRel() == null
										&& getProps().getShortAnnoPrimRel() == null) {
									// There is only one primary text so we can
									// safely assume this is the one
									// the annotation is connected to.
									setAnnotationPrimCouple(
											primaryTextTierList.get(0),
											annoPrimRelations, currColumn,
											headerRow);
                                    primWasFound = true;
								}
							}
                            if(!primWasFound) {
                                annosWithoutPrim.add(tierName);
                            }
							currColumn++;
						}
					}
				}
                
				final Map<String, SLayer> layerTierCouples = getLayerTierCouples();
				Table<Integer, Integer, CellRangeAddress> mergedCells = null;
				if(corpusSheet.getNumMergedRegions() > 0){
				mergedCells = calculateMergedCellIndex(corpusSheet
						.getMergedRegions());
				} 
				int progressTotalNumberOfColumns = 0;
				if (!primTextPos.isEmpty()) {
					progressTotalNumberOfColumns = setPrimText(corpusSheet,
							primTextPos, annoPrimRelations, headerRow,
							mergedCells, layerTierCouples);
				} else {
					SpreadsheetImporter.logger
							.warn("No primary text for the document \""
									+ getResourceURI().lastSegment()
									+ "\" found. Please check the spelling of your properties.");
				}

				setAnnotations(annoPrimRelations, corpusSheet, mergedCells,
						layerTierCouples, progressTotalNumberOfColumns);
			}
			if (getProps().getMetaAnnotation()) {
				setDocMetaData(workbook);
			}
            
            // report if any column was not included
            if(!annosWithoutPrim.isEmpty()) {
                SpreadsheetImporter.logger.warn("No primary text column found for columns\n- {}\nin document {}. This means these columns are not included in the conversion!", 
                        Joiner.on("\n- ").join(annosWithoutPrim), getResourceURI().toFileString());
            }
		}
	}

	/**
	 * Add annotations to the salt graph
	 * 
	 * @param annoPrimRelations
	 * @param corpusSheet
	 * @param mergedCells
	 * @param layerTierCouples
	 * @param progressProcessedNumberOfColumns
	 */
	private void setAnnotations(HashMap<Integer, Integer> annoPrimRelations,
			Sheet corpusSheet,
			Table<Integer, Integer, CellRangeAddress> mergedCells,
			Map<String, SLayer> layerTierCouples,
			int progressProcessedNumberOfColumns) {
		if (!annoPrimRelations.isEmpty()) {
			Row headerRow = corpusSheet.getRow(0);
			DataFormatter formatter = new DataFormatter();
			int progressTotalNumberOfColumns = annoPrimRelations.keySet()
					.size();
			for (int annoTier : annoPrimRelations.keySet()) {

				SSpan annoSpan = null;
				int currAnno = 1;

				while (currAnno < corpusSheet.getPhysicalNumberOfRows()) {
					String annoName = headerRow.getCell(annoTier).toString();
					Row row = corpusSheet.getRow(currAnno);
					Cell annoCell = row.getCell(annoTier);

					if (annoCell != null && !annoCell.toString().isEmpty()) {
						String annoText = "";
						annoText = formatter.formatCellValue(annoCell);

						int annoStart = currAnno - 1;
						int annoEnd = getLastCell(annoCell, mergedCells);
						DataSourceSequence<Integer> sequence = new DataSourceSequence<Integer>();
						sequence.setStart(annoStart);
						sequence.setEnd(annoEnd);
						sequence.setDataSource(getDocument().getDocumentGraph()
								.getTimeline());

						List<SToken> sTokens = getDocument().getDocumentGraph()
								.getTokensBySequence(sequence);

						List<SToken> tokenOfSpan = new ArrayList<>();

						if (sTokens == null) {
							SpreadsheetImporter.logger
									.error("Segmentation error: The segmentation of the tier \""
											+ headerRow.getCell(annoTier)
													.toString()
											+ "\" in the document: \""
											+ getResourceURI().lastSegment()
											+ "\" in line: "
											+ currAnno
											+ " does not match to its primary text: \""
											+ headerRow.getCell(
													annoPrimRelations
															.get(annoTier))
													.toString() + "\".");
						} else {
							for (SToken tok : sTokens) {
								STextualDS textualDS = getTextualDSForNode(tok,
										getDocument().getDocumentGraph());
								if (textualDS.getName().equals(
										headerRow
												.getCell(
														annoPrimRelations
																.get(annoTier))
												.toString())) {
									tokenOfSpan.add(tok);
								}
							}
						}

						annoSpan = getDocument().getDocumentGraph().createSpan(
								tokenOfSpan);

						if (annoSpan != null && annoName != null
								&& !annoName.isEmpty()) {
							// remove primary text info of annotation if given
							if (annoName.matches(".+\\[.+\\]")) {
								annoName = annoName.split("\\[")[0];
							}
							annoSpan.createAnnotation(null, annoName, annoText);
							annoSpan.setName(annoName);
						}
					}

					if (getProps().getLayer() != null && annoSpan != null) {

						if (layerTierCouples.size() > 0) {
							if (layerTierCouples.get(annoName) != null) {
								SLayer sLayer = layerTierCouples.get(annoName);
								getDocument().getDocumentGraph().addLayer(
										sLayer);
								sLayer.addNode(annoSpan);
							}
						}
					}
					currAnno++;
				} // end for each row of annotation

				progressProcessedNumberOfColumns++;
        if(progressProcessedNumberOfColumns < progressTotalNumberOfColumns) {
          setProgress((double) progressProcessedNumberOfColumns
              / (double) progressTotalNumberOfColumns);
        } else {
          setProgress(1.0);
        }
			} // end for each annotation layer
		} else {
			SpreadsheetImporter.logger
					.warn("No annotations except for primary texts found in document \""
							+ getResourceURI().lastSegment() + "\".");
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
	 * @param mergedCells
	 * @param layerTierCouples
	 * @return
	 */
	private int setPrimText(Sheet corpusSheet, List<Integer> primTextPos,
			HashMap<Integer, Integer> annoPrimRelations, Row headerRow,
			Table<Integer, Integer, CellRangeAddress> mergedCells,
			Map<String, SLayer> layerTierCouples) {
		// initialize with number of token we have to create
		int progressTotalNumberOfColumns = primTextPos.size();
		// add each annotation to this number

		int progressProcessedNumberOfColumns = 0;

		// use formater to ensure that e.g. integers will not be converted into
		// decimals
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

				} else if (getProps().getIncludeEmptyPrimCells()) {
					text = "";

				}
				if (text != null) {
					int start = offset;
					int end = start + text.length();
					offset += text.length();
					currentText.append(text);

					currTok = getDocument().getDocumentGraph().createToken(
							primaryText, start, end);

					if (primCell != null) {
						endCell = getLastCell(primCell, mergedCells);
					}
				}

				if (currTok != null) {
					if (lastTok != null && getProps().getAddOrderRelation()) {
						addOrderRelation(lastTok, currTok,
								headerRow.getCell(primText).toString());
					}
					// add timeline relation
					addTimelineRelation(currTok, currRow, endCell, corpusSheet);

					// remember all SToken
					currentTokList.add(currTok);

					// insert space between tokens
					if (text != null
							&& (currRow != corpusSheet.getLastRowNum())) {
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
      if(progressProcessedNumberOfColumns < progressTotalNumberOfColumns) {
        setProgress((double) progressProcessedNumberOfColumns
            / (double) progressTotalNumberOfColumns);
      } else {
        setProgress(1.0);
      }

			if (getProps().getLayer() != null) {
				if (currentTokList != null && layerTierCouples.size() > 0) {
					if (layerTierCouples.get(primaryText.getName()) != null) {
						SLayer sLayer = layerTierCouples.get(primaryText
								.getName());
						getDocument().getDocumentGraph().addLayer(sLayer);
						for (SToken t : currentTokList) {
							sLayer.addNode(t);
						}
					}
				}
			}
		} // end for each primTextPos
		return progressProcessedNumberOfColumns;
	}

	private void addTimelineRelation(SToken tok, int currRow, int endTime,
			Sheet corpusSheet) {
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
	private void setAnnotationPrimCouple(String primTier,
			HashMap<Integer, Integer> annoPrimRelations, int currColumn,
			Row headerRow) {

		if (getColumn(primTier, headerRow) != null) {
			int primData = getColumn(primTier, headerRow);
			if (annoPrimRelations.get(currColumn) != null) {
				SpreadsheetImporter.logger
						.warn("The annotation \""
								+ headerRow.getCell(currColumn).toString()
										.split("\\[")[0]
								+ "\" was allready referenced to the primary text in column \""
								+ annoPrimRelations.get(currColumn)
								+ "\". This reference was overwritten. The new primary text for \""
								+ headerRow.getCell(currColumn).toString()
										.split("\\[")[0] + "\" is \""
								+ headerRow.getCell(primData) + "\".");
				annoPrimRelations.remove(currColumn);
				annoPrimRelations.put(currColumn, primData);
			} else {
				annoPrimRelations.put(currColumn, primData);
			}
		} else {
			SpreadsheetImporter.logger.warn("The primary text \"" + primTier
					+ "\" does not exist in the document \""
					+ getResourceURI().lastSegment() + "\".");
		}
	}

	/**
	 * Return the last cell of merged cells
	 * 
	 * @param primCell
	 *            , current cell of the primary text
	 * @return
	 */
	private int getLastCell(Cell cell,
			Table<Integer, Integer, CellRangeAddress> mergedCellsIdx) {
		int lastCell = cell.getRowIndex();
		CellRangeAddress mergedCell = null;
		if(mergedCellsIdx != null){
		mergedCell = mergedCellsIdx.get(cell.getRowIndex(),
				cell.getColumnIndex());
		}
		if (mergedCell != null) {
			lastCell = mergedCell.getLastRow();
		}
		return lastCell;
	}

	/**
	 * get the primary text of the annotation tier
	 * 
	 * @param currentTier
	 */
	private String getPrimOfAnnoPrimRel(String currentTier) {
		String annoPrimNew = null;

		// check if both properties where used
		if (getProps().getAnnoPrimRel() != null
				&& getProps().getShortAnnoPrimRel() != null) {
			SpreadsheetImporter.logger
					.error("Wrong property handling. Please use only one property to specify which annotation refers to which primary text tier (exclusive use of either 'annoPrimRel' or 'shortAnnoPrimRel').");
		} else {
			String annoPrimRel = getProps().getAnnoPrimRel();
			if (annoPrimRel != null) {
				if (annoPrimRel.matches("(?s).+],.+") || annoPrimRel.matches("(?s).+]")) {
					List<String> annoPrimRelation = Arrays.asList(annoPrimRel
							.split("\\s*,\\s*"));
					for (String annoPrim : annoPrimRelation) {
						String[] splitted = annoPrim.split("=", 2);
						if (splitted.length > 1) {
							String annoName = splitted[0];
							String annoPrimCouple = splitted[1];

							if (annoName.equals(currentTier)) {
								annoPrimNew = annoPrimCouple.split("\\[")[1]
										.replace("]", "");
							}
						} else {
							SpreadsheetImporter.logger
									.error("Can not match the annotations to their primary text, because of syntax errors. "
                    + "Please check the syntax of your property settings. "
                    + "The problematic entry is:\n" 
                    + annoPrim);
						}
					}
				} else {
					SpreadsheetImporter.logger
							.error("Can not match the annotations to their primary text, because of syntax errors. "
                + "Please check the syntax of your property settings.");
				}
			}
			String shortAnnoPrimRel = getProps().getShortAnnoPrimRel();
			if (shortAnnoPrimRel != null) {
				if (shortAnnoPrimRel.matches("(?s).+},.+")
						|| shortAnnoPrimRel.matches("(?s).+}")) {
					List<String> shortAnnoPrimRelation = Arrays
							.asList(shortAnnoPrimRel.split("\\s*},\\s*"));
					for (String annoPrim : shortAnnoPrimRelation) {
						List<String> primAnnoPair = Arrays.asList(annoPrim
								.split("="));
						if (primAnnoPair.size() == 2) {
							String[] annos = primAnnoPair.get(1).split(
									"\\s*,\\s*");
							for (String anno : annos) {
								if (anno.replace("}", "").replace("{", "")
										.equals(currentTier)) {
									annoPrimNew = primAnnoPair.get(0);
								}
							}
						} else {
							SpreadsheetImporter.logger
									.error("Can not match the annotations to their primary text because the property settings do not match the needed syntax (missing \"=\"). Please check the syntax of your property settings.");
						}
					}
				} else {
					SpreadsheetImporter.logger
							.error("Can not match the annotations to their primary text, because of syntax errors. Please check the syntax of your property settings.");
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

			if (row.getCell(currCol) != null
					&& row.getCell(currCol).toString().equals(tierName)) {
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
	public static STextualDS getTextualDSForNode(SNode node,
			SDocumentGraph graph) {
		if (node != null) {
			List<DataSourceSequence> dataSources = graph
					.getOverlappedDataSourceSequence(node,
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
			if (tierLayerCouple.matches("(?s).+},.+")
					|| tierLayerCouple.matches("(?s).+}")) {
				List<String> annoLayerCoupleList = Arrays
						.asList(tierLayerCouple.split("\\s*},\\s*"));
				for (String annoLayer : annoLayerCoupleList) {
					List<String> annoLayerPair = Arrays.asList(annoLayer
							.split("="));
					if ((annoLayerPair.size() == 2)) {
						SLayer sLayer = SaltFactory.createSLayer();
						sLayer.setName(annoLayerPair.get(0));
						String[] assoAnno = annoLayerPair.get(1).split(
								"\\s*,\\s*");
						for (String tier : assoAnno) {
							tier = tier.replace("{", "");
							tier = tier.replace("}", "");
							annoLayerCoupleMap.put(tier, sLayer);
						}
					} else {
						SpreadsheetImporter.logger
								.error("Can not create SLayer because the property settings do not match the needed syntax (missing \"=\"). Please check the syntax of your property settings.");

					}
				}
			} else {
				SpreadsheetImporter.logger
						.error("Can not create SLayer because the property settings do not match the needed syntax. Please check the syntax of your property settings.");
			}
		}
		return annoLayerCoupleMap;
	}

	private void setDocMetaData(Workbook workbook) {
		Sheet metaSheet = null;

		// default ("Tabelle2"/ second sheet)
		if (getProps().getMetaSheet().equals("Tabelle2")) {
			if (workbook.getNumberOfSheets() > 1) {
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

				if (metaKey != null && !metaKey.toString().isEmpty()) {
					if (metaValue != null && !metaValue.toString().isEmpty()) {
						if (getDocument().getMetaAnnotation(metaKey.toString()) == null) {
							getDocument().createMetaAnnotation(null,
									formatter.formatCellValue(metaKey),
									formatter.formatCellValue(metaValue));
						} else {
							SpreadsheetImporter.logger
									.warn("A meta information with the name \""
											+ formatter
													.formatCellValue(metaKey)
											+ "\" allready exists and will not be replaced.");
						}
					} else {
						SpreadsheetImporter.logger
								.warn("No value for the meta data: \""
										+ metaKey.toString() + "\" found.");
					}
				} else {
					if (metaValue != null && !metaValue.toString().isEmpty()) {
						SpreadsheetImporter.logger
								.warn("No meta annotation name for the value \""
										+ metaValue.toString() + "\" found.");
					}
				}
				currRow++;
			}
		}
	}

}
