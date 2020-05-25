package org.corpus_tools.peppermodules.spreadsheet;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

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
			for (STextualDS ds : graph.getTextualDSs()) {
				int rowIx = 0;
				String colName = ds.getName();
				List<SToken> dsTokens = graph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd()));
				sheet.getRow(rowIx).getCell(colIx).setCellValue(colName);
				for (SToken sTok : graph.getSortedTokenByText(dsTokens)) {
					rowIx += 1;
					Optional<SRelation> optRelation = sTok.getOutRelations().stream().filter((SRelation r) -> r instanceof STimelineRelation).findFirst();
					sheet.getRow(rowIx).getCell(colIx).setCellValue(graph.getText(sTok));
					int height = optRelation.isPresent()? ((STimelineRelation) optRelation.get()).getEnd() - ((STimelineRelation) optRelation.get()).getStart() : 1; 
					if (height > 1) {
						sheet.addMergedRegion(new CellRangeAddress(rowIx, rowIx + height - 1, colIx, colIx));
						rowIx += height - 1;
					}
				}
				colIx += 1;
			}
		}
	}
	
	private void mapAnnotations() {
		
	}
	
	private void mapRelations() {
		
	}
}
