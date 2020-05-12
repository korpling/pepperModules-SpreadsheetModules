package org.corpus_tools.peppermodules.spreadsheet.tests;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.corpus_tools.peppermodules.spreadsheet.Spreadsheet2SaltMapper;
import org.corpus_tools.salt.SALT_TYPE;
import org.corpus_tools.salt.SaltFactory;
import org.corpus_tools.salt.common.SDocumentGraph;
import org.corpus_tools.salt.common.SPointingRelation;
import org.corpus_tools.salt.common.STextualDS;
import org.corpus_tools.salt.common.STextualRelation;
import org.corpus_tools.salt.common.STimeline;
import org.corpus_tools.salt.common.STimelineRelation;
import org.corpus_tools.salt.common.SToken;

public class Salt2SpreadsheetMapperTest {
	
	private Spreadsheet2SaltMapper fixture = null;
	
	private Spreadsheet2SaltMapper getFixture() {
		return this.fixture;
	}
	
	private static class SourceGraph {		
		private static SDocumentGraph instance;
		private static final String[] TOKENS = {"we", "don't", "need", "no", "education", ".", "We", "ain't", "gonna", "go", "there", "."};
		private static final int[][] TIME_VALS = {{0, 1}, {1, 3}, {3, 4}, {4, 5}, {5, 6}, {6, 7}, {7, 8}, {8, 10}, {10, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] TOKENS_A = {"we", "do", "n't", "need", "any", "education", ".", "We", "are", "not", "going", "to", "go", "there", "."};
		private static final int[][] TIME_VALS_A = {{0, 1}, {1, 2}, {2, 3}, {3, 4}, {4, 5}, {5, 6}, {6, 7}, {7, 8}, {8, 9}, {9, 10}, {10, 11}, {11, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] LEMMA_A = {"we", "do", "not", "need", "any", "education", ".", "we", "be", "not", "go", "to", "go", "there", "."};
		private static final String[] POS_A = {"PRON", "AUX", "PART", "VERB", "DET", "NOUN", "PUNCT", "PRON", "AUX", "PART", "VERB", "PART", "VERB", "ADV", "PUNCT"};
		private static final Integer[] HEADS_A = {3, 0, 3, 3, 7, 5, null, 10, 0, 10, 10, 12, 13, 14, null};
		private static final String[] DEPRELS_A = {"subj", "root", "mod", "comp:aux", "det", "comp:obj", null, "subj", "root", "mod", "comp:aux", "udep", "comp:obj", "comp:obj", null};
		private static final String[] TOKENS_B = {"we", "do", "not", "need", "education", ".", "We", "will not", "go", "there", "."};
		private static final int[][] TIME_VALS_B = {{0, 1}, {1, 2}, {2, 3}, {3, 4}, {4, 5}, {5, 6}, {6, 7}, {7, 8}, {8, 12}, {12, 13}, {13, 14}, {14, 15}};
		private static final String[] LEMMA_B = {"we", "do", "not", "need", "education", ".", "we", "will", "go", "there", "."};
		private static final String[] POS_B = {"PRON", "AUX", "PART", "VERB", "NOUN", "PUNCT", "PRON", "AUX", "VERB", "ADV", "PUNCT"};
		private static final String LEMMA_NAME = "lemma";
		private static final String POS_NAME = "pos";
		private static final String DEP_TYPE = "dependency";
		private static final String DEP_NAME = "rel";
		protected SourceGraph() {
			if (instance == null) {
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
				for (int k = 0; k < 3; k++) {
					String[] tokArr = tokenArrays.get(k);
					int[][] timeVals = timeArrays.get(k);
					STextualDS ds = docGraph.createTextualDS(StringUtils.join(tokArr));
					int p = 0;
					for (int i = 0; i < tokArr.length; i++) {
						String token = tokArr[i];
						int start = timeVals[i][0];
						int end = timeVals[i][1];
						SToken tok = SaltFactory.createSToken();
						STextualRelation txtRel = (STextualRelation) docGraph.createRelation(tok, ds, SALT_TYPE.STEXTUAL_RELATION, null);
						txtRel.setStart(p);
						txtRel.setEnd(p + token.length());
						p = txtRel.getEnd();
						STimelineRelation timeRel = (STimelineRelation) docGraph.createRelation(tok, timeline, SALT_TYPE.STIMELINE_RELATION, null);
						timeline.increasePointOfTime(end - timeline.getEnd());
						timeRel.setStart(start);
						timeRel.setEnd(end);
						if (k > 0) {
							String pos_val = posAnnos.get(k - 1)[i];
							String lemma_val = lemmaAnnos.get(k - 1)[i];
							tok.createAnnotation(null, POS_NAME, pos_val);
							tok.createAnnotation(null, LEMMA_NAME, lemma_val);
						}
					}
					if (k == 1) {
						// Dependency annotation
						List<SToken> tokenList = docGraph.getSortedTokenByText(docGraph.getOverlappedTokens(ds, SALT_TYPE.STEXTUAL_RELATION));
						for (int i = 0; i < DEPRELS_A.length; i++) {
							int head_id = HEADS_A[i] - 2;  // provided indices are Excel row indices
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
				}
				instance = docGraph;
			}
		}		
		protected SDocumentGraph getInstance() {
			return instance;
		}
	}
}
