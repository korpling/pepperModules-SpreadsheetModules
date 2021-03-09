package org.corpus_tools.peppermodules.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.pepper.modules.exceptions.PepperModuleDataException;
import org.corpus_tools.pepper.modules.exceptions.PepperModuleException;
import org.corpus_tools.peppermodules.spreadsheet.SpreadsheetExporterProperties.AlignmentValue;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.SOrderRelation;
import org.corpus_tools.salt.common.SSpan;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STextualRelation;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SAnnotation;
import org.corpus_tools.salt.core.SNode;
import org.corpus_tools.salt.core.SRelation;
import org.corpus_tools.salt.util.DataSourceSequence;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Salt2SpreadsheetMapper extends PepperMapperImpl implements PepperMapper {
	private static final String DEFAULT_TOK_NAME = "TOK";
	private static final String ERR_MSG_NO_VALUE = "No value provided for cell. This might be due to a non-specified text or annotation value.";
	private static final String ERR_MSG_NO_TEXTUAL_RELATION = "Token has no textual relation and cannot be mapped.";
	private static final String WARNING_NO_TEXT_ANNOTATION = "No text value has been annotated for token span. Overlapped text is used instead.";
	
	private static final Logger logger = LoggerFactory.getLogger(Salt2SpreadsheetMapper.class);
	
	public Salt2SpreadsheetMapper() {
		super();
		annoQNameToColIx = new HashMap<>();
	}
	
	@Override
	public DOCUMENT_STATUS mapSDocument() {
		SDocument document = getDocument();
		if (document == null) {
			throw new PepperModuleDataException(this, "No document provided to mapper.");
		}
		if (document.getDocumentGraph() == null) {
			throw new PepperModuleDataException(this, "Provided document has no document graph.");
		}
		readProperties();
		mapTokenizations();
		mapSpansAndAnnotations();
		writeWorkbook();
		return DOCUMENT_STATUS.COMPLETED;
	}
	
	private Map<String, Integer> columnOrder = null;
	private boolean trimValues = false;
	private Set<String> ignoreNames = null;
	private CellStyle cellStyle = null;
	
	private void readProperties() {
		SpreadsheetExporterProperties properties = (SpreadsheetExporterProperties) getProperties();		
		columnOrder = properties.getColumnOrder();
		trimValues = properties.trimValues();
		ignoreNames = properties.ignoreAnnoNames();
		cellStyle = getWorkbook().createCellStyle();
		{
			String fontName = properties.getFont();
			if (fontName != null) {
				Font defaultFont = getWorkbook().createFont();
				defaultFont.setFontName(fontName);
				cellStyle.setFont(defaultFont);
			}
			Map<AlignmentValue, Short> alignmentMap; {
				alignmentMap = new HashMap<>();
				alignmentMap.put(AlignmentValue.bottom, CellStyle.VERTICAL_BOTTOM);
				alignmentMap.put(AlignmentValue.mid, CellStyle.VERTICAL_CENTER);
				alignmentMap.put(AlignmentValue.top, CellStyle.VERTICAL_TOP);
				alignmentMap.put(AlignmentValue.left, CellStyle.ALIGN_LEFT);
				alignmentMap.put(AlignmentValue.center, CellStyle.ALIGN_CENTER);
				alignmentMap.put(AlignmentValue.right, CellStyle.ALIGN_RIGHT);
			}
			cellStyle.setAlignment( alignmentMap.get(properties.getHorizontalTextAlignment()) );
			cellStyle.setVerticalAlignment( alignmentMap.get(properties.getVerticalTextAlignment()) );
			cellStyle.setDataFormat( getWorkbook().createDataFormat().getFormat("@") );
		}
	}
	
	private Workbook workbook = null;
	
	private Workbook getWorkbook() {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
			workbook.createSheet(getDocument().getName());
		}
		return workbook;
	}
	
	private Sheet getSheet() {
		return getWorkbook().getSheetAt(0);
	}
	
	private SDocumentGraph getDocumentGraph() {
		return getDocument().getDocumentGraph();
	}
	
	private Map<SToken, int[]> tokToCoords = null;
	private Map<String, Integer> annoQNameToColIx = null;
	
	private void mapTokenizations() {
		tokToCoords = new HashMap<SToken, int[]>();
		SDocumentGraph graph = getDocumentGraph();
		List<SRelation<?, ?>> timelineRelations = graph.getRelations().stream()
				.filter((SRelation<?, ?> r) -> r instanceof STimelineRelation)
				.collect(Collectors.toList());
		List<SRelation<?, ?>> orderRelations = graph.getRelations().stream()
				.filter((SRelation<?, ?> r) -> r instanceof SOrderRelation)
				.collect(Collectors.toList());
		Set<String> orderNames = orderRelations.stream().map(SRelation::getType).collect(Collectors.toSet());		
		int headerOffset = 1; // FIXME this appears to be a bit oversimplified
		if (timelineRelations.isEmpty() && orderNames.size() < 2) {
			createEntry(0, 0, 1, DEFAULT_TOK_NAME);
			int rowIx = 0;
			for (STextualDS ds : graph.getTextualDSs()) {
				List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));				
				for (SToken sTok : graph.getSortedTokenByText(dsTokens)) {
					tokToCoords.put(sTok, createEntry(rowIx++, 0, 1, graph.getText(sTok)));
				}
			}
			annoQNameToColIx.put(null, 0);
		} else {
			// multiple parallel tokenizations, go by (1) Timeline relations if possible if not by order relations
			if (timelineRelations.size() > 0) { // by timeline relations is prioritized over ordered spans
				int colIx = 0;
				for (STextualDS ds : graph.getTextualDSs()) {
					String colName = ds.getName();
					List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));
					createEntry(0, colIx, 1, colName);
					for (SToken sTok : graph.getSortedTokenByText(dsTokens)) {
						String text = graph.getText(sTok);
						Optional<SRelation> optRelation = sTok.getOutRelations().stream().filter((SRelation r) -> r instanceof STimelineRelation).findFirst();
						if (!optRelation.isPresent()) {
							throw new PepperModuleDataException(this, "Token has no timeline relation and cannot be mapped: " + text + " (" + colName + ")");
						}
						STimelineRelation tRel = (STimelineRelation) optRelation.get();											 
						tokToCoords.put(sTok, createEntry(tRel.getStart() + headerOffset, colIx, tRel.getEnd() - tRel.getStart(), text));
					}
					colIx += 1;
				}
				annoQNameToColIx.put(null, graph.getTextualDSs().size() - 1);
			} else { // by order relations
				int colIx = 0;
				int upperBound = getDocumentGraph().getTextualDSs().get(0).getText().length() + 1;
				List<Integer>[] startValues = new ArrayList[orderNames.size()];
				List<Integer>[] endValues = new ArrayList[orderNames.size()];
				SNode[][] nodesByCoords = new SNode[upperBound][orderNames.size()];
				String[] colNames = new String[orderNames.size()];
				for (Entry<String, Iterator<SNode>> entry : getTokenGroups(orderNames).entrySet()) {
					String colName = entry.getKey();
					colNames[colIx] = colName;
					Iterator<SNode> nodes = entry.getValue();
					createEntry(0, colIx, 1, colName);
					startValues[colIx] = new ArrayList<Integer>();
					endValues[colIx] = new ArrayList<Integer>();
					while (nodes.hasNext()) {
						SNode node = nodes.next();
						List<SToken> overlappedTokens = graph.getSortedTokenByText( graph.getOverlappedTokens(node) );
						int start; {
							Optional<SRelation> optRelation = overlappedTokens.get(0).getOutRelations().stream()
									.filter((SRelation r) -> r instanceof STextualRelation).findAny();
							if (!optRelation.isPresent()) {
								throw new PepperModuleDataException(this, ERR_MSG_NO_TEXTUAL_RELATION);
							}
							STextualRelation tRel = (STextualRelation) optRelation.get();
							start = tRel.getStart();
						}
						startValues[colIx].add(start);
						int end; {
							Optional<SRelation> optRelation = overlappedTokens.get(overlappedTokens.size() - 1).getOutRelations().stream()
									.filter((SRelation r) -> r instanceof STextualRelation).findAny();
							if (!optRelation.isPresent()) {
								throw new PepperModuleDataException(this, ERR_MSG_NO_TEXTUAL_RELATION);
							}
							STextualRelation tRel = (STextualRelation) optRelation.get();
							end = tRel.getEnd();							
						}
						endValues[colIx].add(end);
						nodesByCoords[start][colIx] = node;				
					}
					colIx += 1;
				};
				List<Integer> orderedValues; {
					Set<Integer> uniqueValues = new HashSet<>();
					Arrays.stream(startValues).forEach(uniqueValues::addAll);
					Arrays.stream(endValues).forEach(uniqueValues::addAll);
					orderedValues = new ArrayList<Integer>(uniqueValues);
					Collections.sort(orderedValues);
				}
				Map<Integer, Integer> coordMap = new HashMap<>();
				for (int i = 0; i < orderedValues.size(); i++) {
					coordMap.put(orderedValues.get(i), i + headerOffset);
				}
				for (int c = 0; c < startValues.length; c++) {
					for (int r = 0; r < startValues[c].size(); r++) {
						int start = startValues[c].get(r);
						int end = endValues[c].get(r);
						SNode sNode = nodesByCoords[start][c];
						if (sNode instanceof SToken) {
							int[] entry = createEntry(coordMap.get(start), c, coordMap.get(end) - coordMap.get(start), getDocumentGraph().getText(sNode));
							tokToCoords.put((SToken) sNode, entry);
						} else {
							String text;
							SAnnotation anno = sNode.getAnnotation(colNames[c]);
							if (anno == null) {
								logger.warn(WARNING_NO_TEXT_ANNOTATION);
								text = getDocumentGraph().getText(sNode);
							} else {
								text = anno.getValue_STEXT();
							}
							int[] entry = createEntry(coordMap.get(start), c, coordMap.get(end) - coordMap.get(start), text);
							for (SToken sTok : getDocumentGraph().getOverlappedTokens(sNode)) {
								tokToCoords.put(sTok, entry);
							}
						}
					}
				}
				annoQNameToColIx.put(null, orderNames.size() - 1);
			}
		}
		int lastUsedIndex = Collections.max(annoQNameToColIx.values());
		for (Entry<String, Integer> entry : columnOrder.entrySet()) {
			createColumn(entry.getValue() + lastUsedIndex + 1, entry.getKey());
		}
		for (Entry<SToken, int[]> entry : tokToCoords.entrySet()) {
			SToken sTok = entry.getKey();
			int[] coords = entry.getValue();
			for (SAnnotation sAnno : sTok.getAnnotations()) {
				String qName = sAnno.getQName();
				Integer colIx = getColumnIndex(qName);
				createEntry(coords[0], colIx, coords[2], sAnno.getValue_STEXT());
			}
		}
	}
	
	private Map<String, Iterator<SNode>> getTokenGroups(Set<String> tokNames) {
		Map<String, SNode> startNodes = getStartNodes(tokNames);
		Map<String, Iterator<SNode>> tokenizations = new HashMap<>();
		for (String name : tokNames) {
			final SNode startNode = startNodes.get(name);
			tokenizations.put(name, new Iterator<SNode>() {
				private final String tokenizationName = name;
				private SNode pointer = null;
				
				@Override
				public boolean hasNext() {
					if (pointer == null) {
						return true;
					}
					Optional<SRelation> next = getOrderRelOpt(pointer);
					return next.isPresent() && next.get().getTarget() != null;
				}

				@Override
				public SNode next() {					
					if (pointer == null) {
						pointer = startNode;
						return startNode;
					}
					pointer = (SNode) getOrderRelOpt(pointer).get().getTarget();
					return pointer;
				}
				
				private Optional<SRelation> getOrderRelOpt(SNode node) {
					return node.getOutRelations().stream()
							.filter((SRelation r) -> r instanceof SOrderRelation && tokenizationName.equals(r.getType()))
							.findAny();
				}
			});			
		}
		return tokenizations;
	}
	
	private Map<String, SNode> getStartNodes(Set<String> orderNames) {
		Map<String, SNode> startNodes = new HashMap<>();
		List<SRelation> orderRels = getDocumentGraph().getRelations().stream()
				.filter((SRelation r) -> r instanceof SOrderRelation)
				.collect(Collectors.toList());
		for (String name : orderNames) {
			startNodes.put(name, recTraceOrigin((SNode) orderRels.stream().filter((SRelation r) -> name.equals(r.getType())).findAny().get().getSource(), name)); 
		}
		return startNodes;
	}
	
	private SNode recTraceOrigin(SNode startNode, String typeName) {
		Optional<SRelation> inRelOpt = startNode.getInRelations().stream().filter((SRelation r) -> r instanceof SOrderRelation && typeName.equals(r.getType())).findAny();
		if (inRelOpt.isPresent() && inRelOpt.get().getSource() != null) {
			return recTraceOrigin(((SOrderRelation) inRelOpt.get()).getSource(), typeName);
		}
		return startNode;
	}
	
	private void createColumn(int index, String name) {
		annoQNameToColIx.put(name, index);
		createEntry(0, index, 1, name);
	}
	
	private int getColumnIndex(String annoQName) {
		Integer colIx = annoQNameToColIx.get(annoQName);
		if (colIx == null) {
			colIx = Collections.max(annoQNameToColIx.values()) + 1;
			createColumn(colIx, annoQName);
		}
		return colIx;
	}
	
	/** 
	 * @param rowIx
	 * @param colIx
	 * @param height
	 * @param value
	 * @return
	 */
	private int[] createEntry(int rowIx, int colIx, int height, String value) {
		if (logger.isDebugEnabled()) {
			logger.debug("Creating entry from row " + rowIx + " in column " + colIx + " for " + height + " rows to insert value \"" + value + "\"");
		}
		if (value == null) {
			throw new PepperModuleDataException(this, ERR_MSG_NO_VALUE);
		}
		for (int ri = 0; ri < height; ri++) {
			Row row = getSheet().getRow(rowIx + ri);
			if (row == null) {
				row = getSheet().createRow(rowIx + ri);
			}
			Cell cell = row.getCell(colIx);
			if (cell == null) {
				cell = row.createCell(colIx);
				cell.setCellStyle(cellStyle);
			}
		}
		getSheet().getRow(rowIx).getCell(colIx).setCellValue(trimValues? value.trim() : value);
		if (height > 1) {
			getSheet().addMergedRegion(new CellRangeAddress(rowIx, rowIx + height - 1, colIx, colIx));
		}
		return new int[]{rowIx, colIx, height};
	}
	
	private void mapSpansAndAnnotations() {
		List<SSpan> spans = getDocumentGraph().getSpans();
		for (SSpan sSpan : spans) {
			List<SToken> overlappedTokens = getDocumentGraph().getOverlappedTokens(sSpan);
			int minRow = Integer.MAX_VALUE;
			int maxRow = 0;
			for (SToken sTok : overlappedTokens) {
				int[] coords = tokToCoords.get(sTok);
				if (coords == null) {
					if (logger.isDebugEnabled()) {
						StringBuilder debugBuilder = new StringBuilder();
						debugBuilder.append("Token is unknown, but has annotations:");						
						for (SAnnotation sAnno : sSpan.getAnnotations()) {
							debugBuilder.append(sAnno.getName());
							debugBuilder.append("=");
							debugBuilder.append(sAnno.getValue_STEXT());
							debugBuilder.append(";");
						}
						logger.debug(debugBuilder.toString());
					}
					throw new PepperModuleException("Token unknown: " + sTok.getId() + ":\"" + getDocumentGraph().getText(sTok) + "\"");
				}				
				minRow = Integer.min(minRow, coords[0]);
				maxRow = Integer.max(maxRow, coords[0] + coords[2]);
			}
			for (SAnnotation sAnno : sSpan.getAnnotations()) {
				if (!ignoreNames.contains(sAnno.getQName())) {
					int colIx = getColumnIndex(sAnno.getQName());
					createEntry(minRow, colIx, maxRow - minRow, sAnno.getValue_STEXT());
				}
			}
		}
	}
	
	private void writeWorkbook() {
		try {
			File outputFile = null;
			if (getResourceURI().toFileString() != null) {
				outputFile = new File(getResourceURI().toFileString());
			} else {
				outputFile = new File(getResourceURI().toString());
			}
			OutputStream outStream = new FileOutputStream(outputFile);
			getWorkbook().write(outStream);
			outStream.close();
		} catch (IOException e) {
			throw new PepperModuleDataException(this, "Could not write workbook to " + getResourceURI().toFileString());
		}
	}
}
