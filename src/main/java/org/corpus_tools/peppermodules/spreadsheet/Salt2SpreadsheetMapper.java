package org.corpus_tools.peppermodules.spreadsheet;

import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.pepper.modules.exceptions.PepperModuleDataException;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.core.SNode;
import org.corpus_tools.salt.core.SRelation;

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
	
	private Workbook getWorkbook(String filePath) {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
			workbook.createSheet(getResourceURI().lastSegment());
		}
		return workbook;
	}
	
	private SDocumentGraph getDocumentGraph() {
		return getDocument().getDocumentGraph();
	}
	
	private void mapTokenizations() {
		SDocumentGraph graph = getDocumentGraph();
		STimeline timeline = graph.getTimeline();
		List<SRelation<?, ?>> timelineRelations = graph.getRelations().stream()
				.filter((SRelation<?, ?> r) -> r instanceof STimelineRelation)
				.collect(Collectors.toList());
		if (timelineRelations.isEmpty()) {
			// concatenate tokenizations
		} else {
			// multiple parallel tokenizations
			
		}
	}
	
	private void mapAnnotations() {
		
	}
	
	private void mapRelations() {
		
	}
}
