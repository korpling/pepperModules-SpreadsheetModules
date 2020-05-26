package org.corpus_tools.peppermodules.spreadsheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
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
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.SSpan;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SAnnotation;
import org.corpus_tools.salt.core.SRelation;
import org.corpus_tools.salt.util.DataSourceSequence;

public class Salt2SpreadsheetMapper extends PepperMapperImpl implements PepperMapper {
	private static final String DEFAULT_TOK_NAME = "TOK";
	
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
		mapTokenizations();
		mapSpansAndAnnotations();
		mapRelations();
		writeWorkbook();
		return DOCUMENT_STATUS.COMPLETED;
	}
	
	private Workbook workbook = null;
	private Sheet sheet = null;
	
	private Workbook getWorkbook() {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
			String ext = getResourceURI().fileExtension();
			String fileName = getResourceURI().lastSegment();
			workbook.createSheet(fileName.substring(0, fileName.length() - ext.length() - 1));
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
		Sheet sheet = getSheet();
		SDocumentGraph graph = getDocumentGraph();
		STimeline timeline = graph.getTimeline();
		List<SRelation<?, ?>> timelineRelations = graph.getRelations().stream()
				.filter((SRelation<?, ?> r) -> r instanceof STimelineRelation)
				.collect(Collectors.toList());
		// FIXME this appears to be a bit oversimplified
		int headerOffset = 1;
		if (timelineRelations.isEmpty()) {
			createEntry(0, 0, 1, DEFAULT_TOK_NAME);
			int rowIx = 0;
			for (STextualDS ds : graph.getTextualDSs()) {
				List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));				
				for (SToken sTok : graph.getSortedTokenByText(dsTokens)) {
					tokToCoords.put(sTok, createEntry(rowIx++, 0, 1, graph.getText(sTok)));
				}
			}
		} else {
			// multiple parallel tokenizations
			int colIx = 0;
			for (STextualDS ds : graph.getTextualDSs()) {
				int rowIx = 0;
				String colName = ds.getName();
				List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));
				createEntry(rowIx, colIx, 1, colName);
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
		}
		annoQNameToColIx.put(null, (int) sheet.getRow(0).getLastCellNum());
		for (Entry<SToken, int[]> entry : tokToCoords.entrySet()) {
			SToken sTok = entry.getKey();
			int[] coords = entry.getValue();
			for (SAnnotation sAnno : sTok.getAnnotations()) {
				String qName = sAnno.getQName();
				Integer colIx = getColumnIndex(qName);
				createEntry(coords[0], coords[1], coords[2], sAnno.getValue_STEXT());
			}
		}
	}
	
	private int getColumnIndex(String annoQName) {
		Integer colIx = annoQNameToColIx.get(annoQName);
		if (colIx == null) {
			colIx = Collections.max(annoQNameToColIx.values()) + 1;
			annoQNameToColIx.put(annoQName, colIx);
			/* create column title */
			createEntry(0, colIx, 1, annoQName);
		}
		return colIx;
	}
	
	private int[] createEntry(int rowIx, int colIx, int height, String value) {
		for (int ri = 0; ri < height; ri++) {
			Row row = getSheet().getRow(rowIx + ri);
			if (row == null) {
				row = getSheet().createRow(rowIx + ri);
			}
			Cell cell = row.getCell(colIx);
			if (cell == null) {
				cell = row.createCell(colIx);
			}
		}
		getSheet().getRow(rowIx).getCell(colIx).setCellValue(value);
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
					throw new PepperModuleException("Token unknown: " + sTok.getId() + ":" + getDocumentGraph().getText(sTok));
				}
				minRow = Integer.min(minRow, coords[0]);
				maxRow = Integer.max(maxRow, coords[1] + coords[2] - 1);
			}
			for (SAnnotation sAnno : sSpan.getAnnotations()) {
				int colIx = 
			}
		}
	}
	
	private void mapRelations() {
		
	}
	
	private void writeWorkbook() {
		try {
			OutputStream outStream = new FileOutputStream(getResourceURI().toFileString());
			getWorkbook().write(outStream);
			outStream.close();
		} catch (IOException e) {
			throw new PepperModuleDataException(this, "Could not write workbook to " + getResourceURI().toFileString());
		}
	}
}
