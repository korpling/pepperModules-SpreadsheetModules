package org.corpus_tools.peppermodules.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
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
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.core.SRelation;
import org.corpus_tools.salt.util.DataSourceSequence;

public class Salt2SpreadsheetMapper extends PepperMapperImpl implements PepperMapper {
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
		mapAnnotations();
		mapRelations();
		writeWorkbook();
		return DOCUMENT_STATUS.COMPLETED;
	}
	
	private Workbook workbook = null;
	private Sheet sheet = null;
	
	private Workbook getWorkbook() {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
			workbook.createSheet(getResourceURI().lastSegment());
		}
		return workbook;
	}
	
	private Sheet getSheet() {
		return getWorkbook().getSheetAt(0);
	}
	
	private SDocumentGraph getDocumentGraph() {
		return getDocument().getDocumentGraph();
	}
	
	private void mapTokenizations() {
		Sheet sheet = getSheet();
		SDocumentGraph graph = getDocumentGraph();
		STimeline timeline = graph.getTimeline();
		List<SRelation<?, ?>> timelineRelations = graph.getRelations().stream()
				.filter((SRelation<?, ?> r) -> r instanceof STimelineRelation)
				.collect(Collectors.toList());
		// FIXME this appears to be a bit oversimplified
		if (timelineRelations.isEmpty()) {
			// concatenate tokenizations
		} else {
			// multiple parallel tokenizations
			int colIx = 0;
			int headerOffset = 1;
			for (STextualDS ds : graph.getTextualDSs()) {
				int rowIx = 0;
				String colName = ds.getName();
				List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));
				createEntry(rowIx, colIx, 1, colName);
				for (SToken sTok : graph.getSortedTokenByText(dsTokens)) {
					rowIx += 1;
					int height = 1;
					Optional<SRelation> optRelation = sTok.getOutRelations().stream().filter((SRelation r) -> r instanceof STimelineRelation).findFirst();
					if (optRelation.isPresent()) {
						STimelineRelation tRel = (STimelineRelation) optRelation.get();
						rowIx = tRel.getStart();
						height = tRel.getEnd() - tRel.getStart();
					} 
					createEntry(rowIx + headerOffset, colIx, height, graph.getText(sTok));					
					rowIx += height - 1;					
				}
				colIx += 1;
			}
		}
	}
	
	private void createEntry(int x, int y, int height, String value) {
		System.out.println("Creating entry " + value + " @ coords " + x + " " + y + " with height " + height);
		for (int xi = 0; xi < height; xi++) {
			Row row = getSheet().getRow(x + xi);
			if (row == null) {
				row = getSheet().createRow(x + xi);
			}
			Cell cell = row.getCell(y);
			if (cell == null) {
				cell = row.createCell(y);
			}
		}
		getSheet().getRow(x).getCell(y).setCellValue(value);
		if (height > 1) {
			getSheet().addMergedRegion(new CellRangeAddress(x, x + height - 1, y, y));
		}
	}
	
	private void mapAnnotations() {
		
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
