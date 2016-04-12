package org.corpus_tools.peppermodules.spreadsheet.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.corpus_tools.pepper.modules.PepperModuleProperty;
import org.corpus_tools.peppermodules.spreadsheet.Spreadsheet2SaltMapper;
import org.corpus_tools.peppermodules.spreadsheet.SpreadsheetImporterProperties;
import org.corpus_tools.salt.SaltFactory;
import org.eclipse.emf.common.util.URI;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.xml.sax.SAXException;

public class Spreadsheet2SaltMapperTest {

	FileOutputStream outStream = null;
	Workbook xlsWb = null;
	Workbook xlsxWb = null;
	File outFile;

	private Spreadsheet2SaltMapper fixture = null;

	public Spreadsheet2SaltMapper getFixture() {
		return fixture;
	}

	public void setFixture(Spreadsheet2SaltMapper fixture) {
		this.fixture = fixture;
	}

	@Before
	public void setUp() throws FileNotFoundException {
		File tmpDir = new File(System.getProperty("java.io.tmpdir")
				+ "/xls2saltTest/");
		tmpDir.mkdirs();
		outFile = new File(tmpDir.getAbsolutePath()
				+ System.currentTimeMillis() + ".xls");
		outStream = new FileOutputStream(outFile);

		setFixture(new Spreadsheet2SaltMapper());
		getFixture().setDocument(SaltFactory.createSDocument());
		getFixture().getDocument().setDocumentGraph(
				SaltFactory.createSDocumentGraph());
		getFixture().getDocument().setName("sampleXSLX");
		getFixture().setProperties(new SpreadsheetImporterProperties());
	}

	@After
	public void tearDown() throws IOException {
		outStream.close();
	}

	private FileOutputStream createFirstXlsSample() throws IOException {

		xlsWb = new HSSFWorkbook();
		Sheet xlsSheet = xlsWb.createSheet("firstXlsSample");
		Row xlsRow1 = xlsSheet.createRow(0);
		Cell xlsCell1 = xlsRow1.createCell(0);
		xlsCell1.setCellValue("tok");
		Cell xlsCell2 = xlsRow1.createCell(1);
		xlsCell2.setCellValue("anno1[tok]");

		Row xlsRow2 = xlsSheet.createRow(1);
		Cell xlsCell21 = xlsRow2.createCell(0);
		xlsCell21.setCellValue("This");
		Cell xlsCell22 = xlsRow2.createCell(1);
		xlsCell22.setCellValue("pron1");

		Row xlsRow3 = xlsSheet.createRow(2);
		Cell xlsCell31 = xlsRow3.createCell(0);
		xlsCell31.setCellValue("is");
		Cell xlsCell32 = xlsRow3.createCell(1);
		xlsCell32.setCellValue("verb1");

		Row xlsRow4 = xlsSheet.createRow(3);
		Cell xlsCell41 = xlsRow4.createCell(0);
		xlsCell41.setCellValue("an");
		Cell xlsCell42 = xlsRow4.createCell(1);
		xlsCell42.setCellValue("art1");

		Row xlsRow5 = xlsSheet.createRow(4);
		Cell xlsCell51 = xlsRow5.createCell(0);
		xlsCell51.setCellValue("example");
		Cell xlsCell52 = xlsRow5.createCell(1);
		xlsCell52.setCellValue("noun1");

		Row xlsRow6 = xlsSheet.createRow(5);
		Cell xlsCell61 = xlsRow6.createCell(0);
		xlsCell61.setCellValue(".");
		Cell xlsCell62 = xlsRow6.createCell(1);
		xlsCell62.setCellValue("punct1");

		xlsWb.write(outStream);

		return outStream;
	}

	private FileOutputStream createFirstXlsxSample() throws IOException {

		xlsxWb = new XSSFWorkbook();
		Sheet xlsSheet = xlsxWb.createSheet("firstXlsxSample");
		Row xlsRow1 = xlsSheet.createRow(0);
		Cell xlsCell1 = xlsRow1.createCell(0);
		xlsCell1.setCellValue("tok");
		Cell xlsCell2 = xlsRow1.createCell(1);
		xlsCell2.setCellValue("anno1[tok]");

		Row xlsRow2 = xlsSheet.createRow(1);
		Cell xlsCell21 = xlsRow2.createCell(0);
		xlsCell21.setCellValue("This");
		Cell xlsCell22 = xlsRow2.createCell(1);
		xlsCell22.setCellValue("pron1");

		Row xlsRow3 = xlsSheet.createRow(2);
		Cell xlsCell31 = xlsRow3.createCell(0);
		xlsCell31.setCellValue("is");
		Cell xlsCell32 = xlsRow3.createCell(1);
		xlsCell32.setCellValue("verb1");

		Row xlsRow4 = xlsSheet.createRow(3);
		Cell xlsCell41 = xlsRow4.createCell(0);
		xlsCell41.setCellValue("an");
		Cell xlsCell42 = xlsRow4.createCell(1);
		xlsCell42.setCellValue("art1");

		Row xlsRow5 = xlsSheet.createRow(4);
		Cell xlsCell51 = xlsRow5.createCell(0);
		xlsCell51.setCellValue("example");
		Cell xlsCell52 = xlsRow5.createCell(1);
		xlsCell52.setCellValue("noun1");

		Row xlsRow6 = xlsSheet.createRow(5);
		Cell xlsCell61 = xlsRow6.createCell(0);
		xlsCell61.setCellValue(".");
		Cell xlsCell62 = xlsRow6.createCell(1);
		xlsCell62.setCellValue("punct1");

		xlsxWb.write(outStream);

		return outStream;
	}

	private FileOutputStream createSecondXlsxSample() throws IOException {

		xlsxWb = new XSSFWorkbook();
		Sheet xlsxSheet = xlsxWb.createSheet("secondXlsxSample");
		Row xlsxRow1 = xlsxSheet.createRow(0);
		Cell xlsxCell1 = xlsxRow1.createCell(0);
		xlsxCell1.setCellValue("tok");
		Cell xlsxCell2 = xlsxRow1.createCell(1);
		xlsxCell2.setCellValue("anno1[tok]");
		Cell xlsxCell3 = xlsxRow1.createCell(2);
		xlsxCell3.setCellValue("tok2");
		Cell xlsxCell4 = xlsxRow1.createCell(3);
		xlsxCell4.setCellValue("anno2[tok2]");

		Row xlsxRow2 = xlsxSheet.createRow(1);
		Cell xlsxCell21 = xlsxRow2.createCell(0);
		xlsxCell21.setCellValue("This");
		Cell xlsxCell22 = xlsxRow2.createCell(1);
		xlsxCell22.setCellValue("pron1");
		Cell xlsxCell23 = xlsxRow2.createCell(2);
		xlsxCell23.setCellValue("This");
		Cell xlsxCell24 = xlsxRow2.createCell(3);
		xlsxCell24.setCellValue("pron2");
		
		Row xlsxRow3 = xlsxSheet.createRow(2);
		Cell xlsxCell31 = xlsxRow3.createCell(0);
		xlsxCell31.setCellValue("is");
		Cell xlsxCell32 = xlsxRow3.createCell(1);
		xlsxCell32.setCellValue("verb1");
		Cell xlsxCell33 = xlsxRow3.createCell(2);
		xlsxCell33.setCellValue("is");
		Cell xlsxCell34 = xlsxRow3.createCell(3);
		xlsxCell34.setCellValue("verb2");

		Row xlsxRow4 = xlsxSheet.createRow(3);
		Cell xlsxCell41 = xlsxRow4.createCell(0);
		xlsxCell41.setCellValue("an");
		Cell xlsxCell42 = xlsxRow4.createCell(1);
		xlsxCell42.setCellValue("art1");
		Cell xlsxCell43 = xlsxRow4.createCell(2);
		xlsxCell43.setCellValue("an");
		Cell xlsxCell44 = xlsxRow4.createCell(3);
		xlsxCell44.setCellValue("art2");

		Row xlsRow5 = xlsxSheet.createRow(4);
		Cell xlsCell51 = xlsRow5.createCell(0);
		xlsCell51.setCellValue("example");
		Cell xlsCell52 = xlsRow5.createCell(1);
		xlsCell52.setCellValue("noun1");
		Cell xlsCell53 = xlsRow5.createCell(2);
		xlsCell53.setCellValue("ex-");
		Cell xlsCell54 = xlsRow5.createCell(3);
		xlsCell54.setCellValue("noun2");

		Row xlsxRow6 = xlsxSheet.createRow(5);
		xlsxRow6.createCell(0);
		xlsxRow6.createCell(1);
		Cell xlsxCell63 = xlsxRow6.createCell(2);
		xlsxCell63.setCellValue("ample");
		xlsxRow6.createCell(3);
		CellRangeAddress mergedPrim = new CellRangeAddress(4, 5, 0, 0);
		xlsxSheet.addMergedRegion(mergedPrim);
		CellRangeAddress mergedAnno1 = new CellRangeAddress(4, 5, 1, 1);
		xlsxSheet.addMergedRegion(mergedAnno1);
		CellRangeAddress mergedAnno2 = new CellRangeAddress(4, 5, 3, 3);
		xlsxSheet.addMergedRegion(mergedAnno2);
		
		Row xlsxRow7 = xlsxSheet.createRow(6);
		Cell xlsxCell71 = xlsxRow7.createCell(0);
		xlsxCell71.setCellValue(".");
		Cell xlsxCell72 = xlsxRow7.createCell(1);
		xlsxCell72.setCellValue("punct1");
		Cell xlsxCell73 = xlsxRow7.createCell(2);
		xlsxCell73.setCellValue(".");
		Cell xlsxCell74 = xlsxRow7.createCell(3);
		xlsxCell74.setCellValue("punct2");

		xlsxWb.write(outStream);

		return outStream;
	}
	
	private FileOutputStream createThirdXlsxSample() throws IOException {

		xlsxWb = new XSSFWorkbook();
		Sheet xlsxSheet = xlsxWb.createSheet("ThirdXlsxSample");
		Row xlsxRow1 = xlsxSheet.createRow(0);
		Cell xlsxCell1 = xlsxRow1.createCell(0);
		xlsxCell1.setCellValue("tok");
		Cell xlsxCell2 = xlsxRow1.createCell(1);
		xlsxCell2.setCellValue("anno1");
		Cell xlsxCell3 = xlsxRow1.createCell(2);
		xlsxCell3.setCellValue("tok2");
		Cell xlsxCell4 = xlsxRow1.createCell(3);
		xlsxCell4.setCellValue("anno2");

		Row xlsxRow2 = xlsxSheet.createRow(1);
		Cell xlsxCell21 = xlsxRow2.createCell(0);
		xlsxCell21.setCellValue("This");
		Cell xlsxCell22 = xlsxRow2.createCell(1);
		xlsxCell22.setCellValue("pron1");
		Cell xlsxCell23 = xlsxRow2.createCell(2);
		xlsxCell23.setCellValue("This");
		Cell xlsxCell24 = xlsxRow2.createCell(3);
		xlsxCell24.setCellValue("pron2");
		
		Row xlsxRow3 = xlsxSheet.createRow(2);
		Cell xlsxCell31 = xlsxRow3.createCell(0);
		xlsxCell31.setCellValue("is");
		Cell xlsxCell32 = xlsxRow3.createCell(1);
		xlsxCell32.setCellValue("verb1");
		Cell xlsxCell33 = xlsxRow3.createCell(2);
		xlsxCell33.setCellValue("is");
		Cell xlsxCell34 = xlsxRow3.createCell(3);
		xlsxCell34.setCellValue("verb2");

		Row xlsxRow4 = xlsxSheet.createRow(3);
		Cell xlsxCell41 = xlsxRow4.createCell(0);
		xlsxCell41.setCellValue("an");
		Cell xlsxCell42 = xlsxRow4.createCell(1);
		xlsxCell42.setCellValue("art1");
		Cell xlsxCell43 = xlsxRow4.createCell(2);
		xlsxCell43.setCellValue("an");
		Cell xlsxCell44 = xlsxRow4.createCell(3);
		xlsxCell44.setCellValue("art2");

		Row xlsRow5 = xlsxSheet.createRow(4);
		Cell xlsCell51 = xlsRow5.createCell(0);
		xlsCell51.setCellValue("ex-Fucking-ample");
		Cell xlsCell52 = xlsRow5.createCell(1);
		xlsCell52.setCellValue("noun1");
		Cell xlsCell53 = xlsRow5.createCell(2);
		xlsCell53.setCellValue("ex-");
		Cell xlsCell54 = xlsRow5.createCell(3);
		xlsCell54.setCellValue("noun2");

		Row xlsxRow6 = xlsxSheet.createRow(5);
		xlsxRow6.createCell(0);
		xlsxRow6.createCell(1);
		Cell xlsxCell63 = xlsxRow6.createCell(2);
		xlsxCell63.setCellValue("ample");
		xlsxRow6.createCell(3);
		CellRangeAddress mergedPrim = new CellRangeAddress(4, 5, 0, 0);
		xlsxSheet.addMergedRegion(mergedPrim);
		CellRangeAddress mergedAnno1 = new CellRangeAddress(4, 5, 1, 1);
		xlsxSheet.addMergedRegion(mergedAnno1);
		CellRangeAddress mergedAnno2 = new CellRangeAddress(4, 5, 3, 3);
		xlsxSheet.addMergedRegion(mergedAnno2);
		
		Row xlsxRow7 = xlsxSheet.createRow(6);
		Cell xlsxCell71 = xlsxRow7.createCell(0);
		xlsxCell71.setCellValue(".");
		Cell xlsxCell72 = xlsxRow7.createCell(1);
		xlsxCell72.setCellValue("punct1");
		Cell xlsxCell73 = xlsxRow7.createCell(2);
		xlsxCell73.setCellValue(".");
		Cell xlsxCell74 = xlsxRow7.createCell(3);
		xlsxCell74.setCellValue("punct2");

		xlsxWb.write(outStream);

		return outStream;
	}
	
	private FileOutputStream createFourthXlsxSample() throws IOException {

		xlsxWb = new XSSFWorkbook();
		Sheet xlsxSheet = xlsxWb.createSheet("fourthXlsxSample");
		Row xlsxRow1 = xlsxSheet.createRow(0);
		Cell xlsxCell1 = xlsxRow1.createCell(0);
		xlsxCell1.setCellValue("tok");
		Cell xlsxCell2 = xlsxRow1.createCell(1);
		xlsxCell2.setCellValue("anno1");
		Cell xlsxCell3 = xlsxRow1.createCell(2);
		xlsxCell3.setCellValue("tok2");
		Cell xlsxCell4 = xlsxRow1.createCell(3);
		xlsxCell4.setCellValue("anno2");
		Cell xlsxCell5 = xlsxRow1.createCell(4);
		xlsxCell5.setCellValue("lb");


		Row xlsxRow2 = xlsxSheet.createRow(1);
		Cell xlsxCell21 = xlsxRow2.createCell(0);
		xlsxCell21.setCellValue("This");
		Cell xlsxCell22 = xlsxRow2.createCell(1);
		xlsxCell22.setCellValue("pron1");
		Cell xlsxCell23 = xlsxRow2.createCell(2);
		xlsxCell23.setCellValue("This");
		Cell xlsxCell24 = xlsxRow2.createCell(3);
		xlsxCell24.setCellValue("pron2");
		Cell xlsxCell25 = xlsxRow2.createCell(4);
		xlsxCell25.setCellValue("lb");
		
		Row xlsxRow3 = xlsxSheet.createRow(2);
		Cell xlsxCell31 = xlsxRow3.createCell(0);
		xlsxCell31.setCellValue("is");
		Cell xlsxCell32 = xlsxRow3.createCell(1);
		xlsxCell32.setCellValue("verb1");
		Cell xlsxCell33 = xlsxRow3.createCell(2);
		xlsxCell33.setCellValue("is");
		Cell xlsxCell34 = xlsxRow3.createCell(3);
		xlsxCell34.setCellValue("verb2");
		xlsxRow3.createCell(4);

		Row xlsxRow4 = xlsxSheet.createRow(3);
		Cell xlsxCell41 = xlsxRow4.createCell(0);
		xlsxCell41.setCellValue("an");
		Cell xlsxCell42 = xlsxRow4.createCell(1);
		xlsxCell42.setCellValue("art1");
		Cell xlsxCell43 = xlsxRow4.createCell(2);
		xlsxCell43.setCellValue("an");
		Cell xlsxCell44 = xlsxRow4.createCell(3);
		xlsxCell44.setCellValue("art2");
		xlsxRow4.createCell(4);

		Row xlsRow5 = xlsxSheet.createRow(4);
		Cell xlsCell51 = xlsRow5.createCell(0);
		xlsCell51.setCellValue("ex-Fucking-ample");
		Cell xlsCell52 = xlsRow5.createCell(1);
		xlsCell52.setCellValue("noun1");
		Cell xlsCell53 = xlsRow5.createCell(2);
		xlsCell53.setCellValue("ex-");
		Cell xlsCell54 = xlsRow5.createCell(3);
		xlsCell54.setCellValue("noun2");
		xlsRow5.createCell(4);

		Row xlsxRow6 = xlsxSheet.createRow(5);
		xlsxRow6.createCell(0);
		xlsxRow6.createCell(1);
		Cell xlsxCell63 = xlsxRow6.createCell(2);
		xlsxCell63.setCellValue("ample");
		xlsxRow6.createCell(3);
		CellRangeAddress mergedPrim = new CellRangeAddress(4, 5, 0, 0);
		xlsxSheet.addMergedRegion(mergedPrim);
		CellRangeAddress mergedAnno1 = new CellRangeAddress(4, 5, 1, 1);
		xlsxSheet.addMergedRegion(mergedAnno1);
		CellRangeAddress mergedAnno2 = new CellRangeAddress(4, 5, 3, 3);
		xlsxSheet.addMergedRegion(mergedAnno2);
		xlsxRow6.createCell(4);
		
		Row xlsxRow7 = xlsxSheet.createRow(6);
		Cell xlsxCell71 = xlsxRow7.createCell(0);
		xlsxCell71.setCellValue(".");
		Cell xlsxCell72 = xlsxRow7.createCell(1);
		xlsxCell72.setCellValue("punct1");
		Cell xlsxCell73 = xlsxRow7.createCell(2);
		xlsxCell73.setCellValue(".");
		Cell xlsxCell74 = xlsxRow7.createCell(3);
		xlsxCell74.setCellValue("punct2");
		xlsxRow7.createCell(4);
		CellRangeAddress mergedAnnoLb = new CellRangeAddress(1, 6, 4, 4);
		xlsxSheet.addMergedRegion(mergedAnnoLb);

		xlsxWb.write(outStream);

		return outStream;
	}
	
	private void start(Spreadsheet2SaltMapper mapper, String xlsString)
			throws FileNotFoundException, UnsupportedEncodingException {

		this.getFixture().setResourceURI(
				URI.createFileURI(outFile.getAbsolutePath()));
		this.getFixture().mapSDocument();
	}

	/**
	 * test for correct generation of primary data.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testPrimData() throws IOException {

		createSecondXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("secondXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);

		start(getFixture(), outStream.toString());
		assertEquals(1, getFixture().getDocument().getDocumentGraph().getTextualDSs().size());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0));
		assertEquals("This is an example .", getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0).getText());
	}

	/**
	 * test for correct generation of token with only one primary text layer in
	 * a xls file.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testPrimDataTokenXls() throws IOException {

		createFirstXlsSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("firstXlsSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(0));
		assertEquals("This", getFixture().getDocument().getDocumentGraph().getText(
								getFixture().getDocument().getDocumentGraph().getTokens().get(0)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(1));
		assertEquals("is", getFixture().getDocument().getDocumentGraph().getText(
								getFixture().getDocument().getDocumentGraph().getTokens().get(1)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(2));
		assertEquals("an", getFixture().getDocument().getDocumentGraph().getText(
								getFixture().getDocument().getDocumentGraph().getTokens().get(2)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(3));
		assertEquals("example", getFixture().getDocument().getDocumentGraph().getText(
								getFixture().getDocument().getDocumentGraph().getTokens().get(3)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(4));
		assertEquals(".", getFixture().getDocument().getDocumentGraph().getText(
								getFixture().getDocument().getDocumentGraph().getTokens().get(4)));
	}

	/**
	 * test for correct generation of token with only one primary text layer in
	 * a xlsx file.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testPrimDataTokenXlsx() throws IOException {

		createSecondXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("secondXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok, tok2");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(0));
		assertEquals("This", getFixture().getDocument().getDocumentGraph()
						.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(0)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(1));
		assertEquals("is", getFixture().getDocument().getDocumentGraph()
						.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(1)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(2));
		assertEquals("an", getFixture().getDocument().getDocumentGraph()
						.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(2)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(3));
		assertEquals("example", getFixture().getDocument().getDocumentGraph()
						.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(3)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(4));
		assertEquals(".", getFixture().getDocument().getDocumentGraph()
						.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(4)));
//		assertEquals("This", getFixture().getDocument().getDocumentGraph()
//				.getText(getFixture().getDocument().getDocumentGraph().getTokens().get(5)));
		assertEquals(11, getFixture().getDocument().getDocumentGraph().getTokens().size());

	}

	/**
	 * test for correct generation of primary text, having multiple primary text layer in
	 * a xlsx file.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testMultiplePrimDataXlsx() throws IOException {
		createSecondXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("secondXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok, tok2");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());
		assertEquals(2, getFixture().getDocument().getDocumentGraph()
				.getTextualDSs().size());
		assertNotNull(getFixture().getDocument().getDocumentGraph()
				.getTextualDSs().get(0));
		assertEquals("This is an example .", getFixture().getDocument()
				.getDocumentGraph().getTextualDSs().get(0).getText());
		assertNotNull(getFixture().getDocument().getDocumentGraph()
				.getTextualDSs().get(1));
		assertEquals("This is an ex- ample .", getFixture().getDocument()
				.getDocumentGraph().getTextualDSs().get(1).getText());
	}
	
	/**
	 * test for correct generation of annotations with concatenated STextualDS
	 * for each primary text (only one span for all annotations).
	 * 
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws IOException
	 * @throws XMLStreamException
	 */
	@Test
	public void testMultiplePrimDataSpan() throws ParserConfigurationException, SAXException, IOException, XMLStreamException {
		// TODO: change to a check for correct spanning regarding to associated annotations
		createSecondXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("secondXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok, tok2");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());

		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(0));
//		assertEquals("tok" , getFixture().getDocument().getDocumentGraph().getSpans().get(0).getName());
//		assertEquals("This is an example .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(0)));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(2));
//		assertEquals("tok2" , getFixture().getDocument().getDocumentGraph().getSpans().get(2).getName());
//		assertEquals("This is an ex- ample .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(2)));
		//		assertEquals(11, getFixture().getDocument().getDocumentGraph().getTimelineRelations().size());
	}
	
	@Test
	public void testMultiplePrimDataOrder() throws IOException {
		createSecondXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET)).setValue("secondXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT)).setValue("tok, tok2");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());
		
		assertNotNull(getFixture().getDocument().getDocumentGraph().getOrderRelations());
		assertEquals(9, getFixture().getDocument().getDocumentGraph().getOrderRelations().size());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(0), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(0).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(1), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(0).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(1), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(1).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(2), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(1).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(2), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(2).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(3), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(2).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(3), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(3).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(4), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(3).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(5), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(4).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(6), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(4).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(6), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(5).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(7), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(5).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(7), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(6).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(8), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(6).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(8), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(7).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(9), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(7).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(9), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(8).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(10), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(8).getTarget());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTimeline());

	}
	
	@Test
	public void testRenameAnnoNames() throws ParserConfigurationException, SAXException, IOException, XMLStreamException {
		// TODO: check for Spans in relation to their annotations rather than their primary text layers.
		createThirdXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("ThirdXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok, tok2");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_ANNO_REFERS_TO))
				.setValue("anno1>anno1[tok], anno2>anno2[tok2]");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());

		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(0));
//		assertEquals("This is an ex-Fucking-ample .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(0)));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(2));
//		assertEquals("This is an ex- ample .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(2)));
		assertNotNull(getFixture().getDocument().getDocumentGraph().getOrderRelations());
		assertEquals(9, getFixture().getDocument().getDocumentGraph().getOrderRelations().size());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(0), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(0).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(1), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(0).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(1), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(1).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(2), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(1).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(2), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(2).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(3), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(2).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(3), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(3).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(4), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(3).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(5), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(4).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(6), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(4).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(6), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(5).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(7), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(5).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(7), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(6).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(8), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(6).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(8), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(7).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(9), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(7).getTarget());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(9), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(8).getSource());
		assertEquals(getFixture().getDocument().getDocumentGraph().getTokens().get(10), getFixture().getDocument().getDocumentGraph().getOrderRelations().get(8).getTarget());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTimeline());
//		assertEquals(11, getFixture().getDocument().getDocumentGraph().getTimelineRelations().size());
	}
	
	@Test
	public void testAnnotations() throws ParserConfigurationException, SAXException, IOException, XMLStreamException {

		createFourthXlsxSample();
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
				.setValue("fourthXlsxSample");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
				.setValue("tok, tok2");
		((PepperModuleProperty<String>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_ANNO_REFERS_TO))
				.setValue("anno1>anno1[tok], anno2>anno2[tok2], lb>lb[tok]");
		((PepperModuleProperty<Boolean>) getFixture().getProperties()
				.getProperty(SpreadsheetImporterProperties.PROP_META_ANNO))
				.setValue(false);
		
		start(getFixture(), outStream.toString());

//		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans());
//		assertEquals(13, getFixture().getDocument().getDocumentGraph().getSpans().size());
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(1));
//		assertEquals("This is an ex-Fucking-ample .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(1)));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getSpans().get(3));
//		assertEquals("This is an ex- ample .", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getSpans().get(3)));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getAnnotations());
//		assertEquals(2, getFixture().getDocument().getDocumentGraph().getTokens().get(0).getAnnotations().size());
//		assertEquals("lb", getFixture().getDocument().getDocumentGraph().getSpans().get(1).getAnnotation("lb").getValue());
//		assertEquals("pron1", getFixture().getDocument().getDocumentGraph().getSpans().get(2).getAnnotation("anno1").getValue());
//		assertEquals("verb1", getFixture().getDocument().getDocumentGraph().getSpans().get(3).getAnnotation("anno1").getValue());
//		assertEquals("art1", getFixture().getDocument().getDocumentGraph().getSpans().get(4).getAnnotation("anno1").getValue());
//		assertEquals("noun1", getFixture().getDocument().getDocumentGraph().getSpans().get(5).getAnnotation("anno1").getValue());
//		assertEquals("punct1", getFixture().getDocument().getDocumentGraph().getSpans().get(6).getAnnotation("anno1").getValue());
//		assertEquals("pron2", getFixture().getDocument().getDocumentGraph().getSpans().get(8).getAnnotation("anno2").getValue());
//		assertEquals("verb2", getFixture().getDocument().getDocumentGraph().getSpans().get(9).getAnnotation("anno2").getValue());
//		assertEquals("art2", getFixture().getDocument().getDocumentGraph().getSpans().get(10).getAnnotation("anno2").getValue());
//		assertEquals("noun2", getFixture().getDocument().getDocumentGraph().getSpans().get(11).getAnnotation("anno2").getValue());
//		assertEquals("punct2", getFixture().getDocument().getDocumentGraph().getSpans().get(12).getAnnotation("anno2").getValue());
	}
	
//	@Test
//	public void testSLayer() throws IOException {
//		createThirdXlsxSample();
//		((PepperModuleProperty<String>) getFixture().getProperties()
//				.getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET))
//				.setValue("firstSheet");
//		((PepperModuleProperty<String>) getFixture().getProperties()
//				.getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT))
//				.setValue("tok, tok2");
//		((PepperModuleProperty<String>) getFixture().getProperties()
//				.getProperty(SpreadsheetImporterProperties.PROP_ANNO_REFERS_TO))
//				.setValue("anno1>anno1[tok], anno2>anno2[tok2]");
//		((PepperModuleProperty<String>) getFixture().getProperties()
//				.getProperty(SpreadsheetImporterProperties.PROP_SET_LAYER))
//				.setValue("textual>{tok, tok2}, morphologigal>{anno1, anno2}, graphical>{lb}");
//		start(getFixture(), outStream.toString());
//
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getLayers());
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getLayerByName("textual"));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getLayerByName("morphologigal"));
//		assertNotNull(getFixture().getDocument().getDocumentGraph().getLayerByName("graphical"));
//		assertEquals(3, getFixture().getDocument().getDocumentGraph().getLayers().size());
//	}
}
