package org.corpus_tools.peppermodules.spreadsheet.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.commons.lang3.tuple.Triple;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.assertj.core.groups.Tuple;
import org.corpus_tools.pepper.testFramework.PepperTestUtil;
import org.corpus_tools.peppermodules.spreadsheet.Salt2SpreadsheetMapper;
import org.corpus_tools.peppermodules.spreadsheet.SpreadsheetExporterProperties;
import org.corpus_tools.salt.SALT_TYPE;
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.SPointingRelation;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STextualRelation;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;
import org.corpus_tools.salt.util.DataSourceSequence;
import org.eclipse.emf.common.util.URI;
import org.junit.Before;
import org.junit.Test;

public class Salt2SpreadsheetMapperTest {
	
	private static final String SAMPLE_FILE_NAME = "exporter_example.xlsx";
	private static final String TEST_OUT_FILE_NAME = "_test.xlsx";
	
	private Salt2SpreadsheetMapper fixture = null;
	
	private Salt2SpreadsheetMapper getFixture() {
		return this.fixture;
	}
	
	@Before
	public void setFixture() {
		this.fixture = new Salt2SpreadsheetMapper();
	}
	
	private Workbook mappingResult = null;
	
	/**
	 * This method runs a conversion of the given {@link SDocumentGraph} Object. 
	 * @param documentGraph The {@link SDocumentGraph} to be exported.
	 * @return {@link Path} of the created excel sheet, null if mapping failed
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	private Workbook map(SDocumentGraph documentGraph) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Salt2SpreadsheetMapper mapper = getFixture();
		Path targetPath = Paths.get(PepperTestUtil.getTempPath_static("exporter_test").toString(), TEST_OUT_FILE_NAME);
		targetPath.toFile().getParentFile().mkdirs();
		mapper.setResourceURI( URI.createFileURI(targetPath.toString()) );
		SpreadsheetExporterProperties properties = new SpreadsheetExporterProperties();
		properties.setPropertyValue(properties.PROP_COL_ORDER, "TOK::doc, TOK::sent, TOK_A::pos, TOK_B::pos, TOK_A::lemma, TOK_B::lemma");
		mapper.setProperties(properties);
		SDocument doc = SaltFactory.createSDocument();
		doc.setDocumentGraph(documentGraph);
		mapper.setDocument(doc);
		mapper.mapSDocument();
		assertTrue("Exporter's output cannot be found on disk.", Files.exists(targetPath));
		return WorkbookFactory.create(targetPath.toFile());
	}
	
	private Workbook getMappingResult() throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (mappingResult == null) {
			SDocumentGraph fixGraph = TestGraph.getGraph();
			mappingResult = map(fixGraph);
		}
		return mappingResult;
	}
	
	private void testColumns(Integer... columns) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook goldWorkbook = WorkbookFactory.create( Paths.get(PepperTestUtil.getTestResources(), SAMPLE_FILE_NAME).toFile() );
		Sheet goldSheet = goldWorkbook.getSheetAt(0);		
		Workbook fixWorkbook = getMappingResult();
		Sheet fixSheet = fixWorkbook.getSheetAt(0);
		// check columns
		assertEquals(goldSheet.getFirstRowNum(), fixSheet.getFirstRowNum());
		assertEquals(goldSheet.getLastRowNum(), fixSheet.getLastRowNum());
		Iterator<Row> itGoldRows = goldSheet.rowIterator();
		Iterator<Row> itFixRows = fixSheet.rowIterator();
		for (Row gRow = itGoldRows.next(), fRow = itFixRows.next(); itGoldRows.hasNext() && itFixRows.hasNext(); gRow = itGoldRows.next(), fRow = itFixRows.next()) {
			for (int c : columns) {
				String gValue = gRow.getCell(c) == null || gRow.getCell(c).getStringCellValue().trim().isEmpty()? null : gRow.getCell(c).getStringCellValue();
				String fValue = fRow.getCell(c) == null || fRow.getCell(c).getStringCellValue().trim().isEmpty()? null : fRow.getCell(c).getStringCellValue();
				assertEquals("Token values are unequal @ " + (fRow.getRowNum() + 1) + " " + c, gValue, fValue);				
			}
		}
		// check for merged cells
		final Set<Integer> colSet = new HashSet<Integer>(Arrays.asList(columns));		
		List<CellRangeAddress> goldMergedRegions = goldSheet.getMergedRegions();
		Set<Tuple> goldRelevantRegions = goldMergedRegions.stream()
				.filter((CellRangeAddress a) -> colSet.contains(new Integer(a.getFirstColumn())))
				.map((CellRangeAddress a) -> Tuple.tuple(a.getFirstRow(), a.getLastRow(), a.getFirstColumn(), a.getLastColumn()))
				.collect(Collectors.toSet());
		List<CellRangeAddress> fixMergedRegions = fixSheet.getMergedRegions();
		Set<Tuple> fixRelevantRegions = fixMergedRegions.stream()
				.filter((CellRangeAddress a) -> colSet.contains(new Integer(a.getFirstColumn())))
				.map((CellRangeAddress a) -> Tuple.tuple(a.getFirstRow(), a.getLastRow(), a.getFirstColumn(), a.getLastColumn()))
				.collect(Collectors.toSet());
		assertEquals(goldRelevantRegions.size(), fixRelevantRegions.size());
		for (Tuple address : goldRelevantRegions) {
			assertTrue(fixRelevantRegions.contains(address));
		}
	}
	
	@Test
	public void testTokenizations() throws EncryptedDocumentException, InvalidFormatException, IOException {
		testColumns(0, 1, 2);
	}
	
	@Test
	public void testTokenAnnotations() throws EncryptedDocumentException, InvalidFormatException, IOException {
		testColumns(5, 6, 7, 8);
	}
	
	@Test
	public void testSpanAnnotations() throws EncryptedDocumentException, InvalidFormatException, IOException {
		testColumns(3, 4);
	}
	
	@Test
	public void testDependencyAnnotations() throws EncryptedDocumentException, InvalidFormatException, IOException {
		testColumns(9, 10);
	}
	
	private static class TestGraph {		
		private static TestGraph instance;
		private static SDocumentGraph instanceGraph;
		private static Workbook workbook;
		private static final String[] TOK_NAMES = {"TOK", "TOK_A", "TOK_B"};
		private static final String[] SENTENCES = {"S1", "S2"};
		private static final int[][] SENTENCE_TOKENS = {{0, 5}, {6, 11}};
		private static final String SENTENCE_NAME = "sent";
		private static final String DOC_ANNO_NAME = "doc";
		private static final String DOC_ANNO_VAL = "Document_1";
		private static final String[] TOKENS = {"we", "don't", "need", "no", "education", ".", "we", "ain't", "gonna", "go", "there", "."};
		private static final int[][] TIME_VALS = {{0, 1}, {1, 3}, {3, 4}, {4, 5}, {5, 6}, {6, 7}, {7, 8}, {8, 10}, {10, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] TOKENS_A = {"we", "do", "n't", "need", "any", "education", ".", "we", "are", "not", "going", "to", "go", "there", "."};
		private static final int[][] TIME_VALS_A = {{0, 1}, {1, 2}, {2, 3}, {3, 4}, {4, 5}, {5, 6}, {6, 7}, {7, 8}, {8, 9}, {9, 10}, {10, 11}, {11, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] LEMMA_A = {"we", "do", "not", "need", "any", "education", ".", "we", "be", "not", "go", "to", "go", "there", "."};
		private static final String[] POS_A = {"PRON", "AUX", "PART", "VERB", "DET", "NOUN", "PUNCT", "PRON", "AUX", "PART", "VERB", "PART", "VERB", "ADV", "PUNCT"};
		private static final Integer[] HEADS_A = {3, 0, 3, 3, 7, 5, null, 10, 0, 10, 10, 12, 13, 14, null};
		private static final String[] DEPRELS_A = {"subj", "root", "mod", "comp:aux", "det", "comp:obj", null, "subj", "root", "mod", "comp:aux", "udep", "comp:obj", "comp:obj", null};
		private static final String[] TOKENS_B = {"we", "do", "not", "need", "education", ".", "we", "will not", "go", "there", "."};
		private static final int[][] TIME_VALS_B = {{0, 1}, {1, 2}, {2, 3}, {3, 4}, {5, 6}, {6, 7}, {7, 8}, {8, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] LEMMA_B = {"we", "do", "not", "need", "education", ".", "we", "will", "go", "there", "."};
		private static final String[] POS_B = {"PRON", "AUX", "PART", "VERB", "NOUN", "PUNCT", "PRON", "AUX", "VERB", "ADV", "PUNCT"};
		private static final String LEMMA_NAME = "lemma";
		private static final String POS_NAME = "pos";
		private static final String DEP_TYPE = "dependency";
		private static final String DEP_NAME = "rel";
		protected static final String DOC_NAME = "test_doc";
		private TestGraph() {			
			SDocumentGraph docGraph = SaltFactory.createSDocumentGraph();
			STimeline timeline = docGraph.createTimeline();
			List<String[]> tokenArrays = new ArrayList<>();
			tokenArrays.add(TOKENS);
			tokenArrays.add(TOKENS_A);
			tokenArrays.add(TOKENS_B);
			List<int[][]> timeArrays = new ArrayList<>();
			timeArrays.add(TIME_VALS);
			timeArrays.add(TIME_VALS_A);
			timeArrays.add(TIME_VALS_B);
			List<String[]> posAnnos = new ArrayList<>();
			posAnnos.add(POS_A);
			posAnnos.add(POS_B);
			List<String[]> lemmaAnnos = new ArrayList<>();
			lemmaAnnos.add(LEMMA_A);
			lemmaAnnos.add(LEMMA_B);
			for (int k = 0; k < 3; k++) { // k iterates over competing tokenizations and their corresponding annotations
				String[] tokArr = tokenArrays.get(k);
				int[][] timeVals = timeArrays.get(k);
				STextualDS ds = docGraph.createTextualDS(StringUtils.join(tokArr));
				ds.setName(TOK_NAMES[k]);
				int p = 0;
				for (int i = 0; i < tokArr.length; i++) { // i iterates over individual values of the k-th tokenization/annotation, i. e. across time
					String token = tokArr[i];
					int start = timeVals[i][0];
					int end = timeVals[i][1];
					SToken tok = SaltFactory.createSToken();	
					docGraph.addNode(tok);
					STextualRelation txtRel = SaltFactory.createSTextualRelation();
					txtRel.setSource(tok);
					txtRel.setTarget(ds);					
					txtRel.setStart(p);
					txtRel.setEnd(p + token.length());
					docGraph.addRelation(txtRel);
					p = txtRel.getEnd();
					STimelineRelation timeRel = SaltFactory.createSTimelineRelation();
					timeRel.setSource(tok);
					timeRel.setTarget(timeline);
					docGraph.addRelation(timeRel);
					int tlLen = timeline.getEnd() == null? 0 : timeline.getEnd();
					timeline.increasePointOfTime(end - tlLen);
					timeRel.setStart(start);
					timeRel.setEnd(end);
					if (k > 0) {
						String pos_val = posAnnos.get(k - 1)[i];
						String lemma_val = lemmaAnnos.get(k - 1)[i];
						tok.createAnnotation(TOK_NAMES[k], POS_NAME, pos_val);
						tok.createAnnotation(TOK_NAMES[k], LEMMA_NAME, lemma_val);
					}
				}
				if (k == 0) {
					List<SToken> tokenList = docGraph.getSortedTokenByText(docGraph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd())));
					for (int si = 0; si < SENTENCES.length; si++) {
						List<SToken> sentenceTokens = tokenList.subList(SENTENCE_TOKENS[si][0], SENTENCE_TOKENS[si][1] + 1);
						docGraph.createSpan(sentenceTokens).createAnnotation(TOK_NAMES[k], SENTENCE_NAME, SENTENCES[si]);
					}
					docGraph.createSpan(tokenList).createAnnotation(TOK_NAMES[k], DOC_ANNO_NAME, DOC_ANNO_VAL);
				}
				else if (k == 1) {
					// Dependency annotation
					List<SToken> tokenList = docGraph.getSortedTokenByText(docGraph.getTokensBySequence(new DataSourceSequence<Number>(ds, ds.getStart(), ds.getEnd())));
					for (int i = 0; i < DEPRELS_A.length; i++) {
						int head_id = HEADS_A[i] == null? -1 : HEADS_A[i] - 2;  // provided indices are Excel row indices
						String depRelLabel = DEPRELS_A[i];
						if (head_id >= 0) {
							SToken sourceTok = tokenList.get(head_id);
							SToken targetTok = tokenList.get(i);
							SPointingRelation depRel = (SPointingRelation) docGraph.createRelation(sourceTok, targetTok, SALT_TYPE.SPOINTING_RELATION, null);
							depRel.setType(DEP_TYPE);
							depRel.createAnnotation(null, DEP_NAME, depRelLabel);
						}
					}
				}				
				instanceGraph = docGraph;
			}
		}		
		
		private static TestGraph getInstance() {
			if (instance == null) {
				instance = new TestGraph();
			}
			return instance;
		}
		
		private SDocumentGraph getInstanceGraph() {
			return instanceGraph;
		}
		
		protected static SDocumentGraph getGraph() {
			return getInstance().getInstanceGraph();
		}
	}
}
