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

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
		File tmpDir = new File(System.getProperty("java.io.tmpdir") + "/xls2saltTest/");
		tmpDir.mkdirs();
		outFile = new File(tmpDir.getAbsolutePath() + System.currentTimeMillis() + ".xls");
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
	
	private FileOutputStream createFirstSample() throws IOException {
		
		xlsWb = new HSSFWorkbook();
		Sheet xlsSheet = xlsWb.createSheet("firstSheet");
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


	private void start(Spreadsheet2SaltMapper mapper, String xlsString)
			throws FileNotFoundException, UnsupportedEncodingException {
		
		this.getFixture().setResourceURI(
				URI.createFileURI(outFile.getAbsolutePath()));
		this.getFixture().mapSDocument();
	}

	/**
	 * test for correct generation of primary data
	 * 
	 * @throws IOException
	 */
	@Test
	public void testPrimData() throws IOException {

		createFirstSample();
		((PepperModuleProperty<String>)getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET)).setValue("firstSheet");
		((PepperModuleProperty<String>)getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT)).setValue("tok");
		
		start(getFixture(), outStream.toString());
		 assertEquals(1,
		 getFixture().getDocument().getDocumentGraph().getTextualDSs().size());
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0));
		 assertEquals("This is an example .", getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0).getText());
	}

	/**
	 * test for correct generation of token with only one primary text layer
	 * 
	 * @throws IOException
	 */
	@Test
	public void testPrimDataToken() throws IOException {

		createFirstSample();
		((PepperModuleProperty<String>)getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_CORPUS_SHEET)).setValue("firstSheet");
		((PepperModuleProperty<String>)getFixture().getProperties().getProperty(SpreadsheetImporterProperties.PROP_PRIMARY_TEXT)).setValue("tok");
		
		start(getFixture(), outStream.toString());
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens());
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(0));
		 assertEquals("This", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getTokens().get(0)));
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(1));
		 assertEquals("is", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getTokens().get(1)));
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(2));
		 assertEquals("an", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getTokens().get(2)));
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(3));
		 assertEquals("example", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getTokens().get(3)));
		 assertNotNull(getFixture().getDocument().getDocumentGraph().getTokens().get(4));
		 assertEquals(".", getFixture().getDocument().getDocumentGraph().getText(getFixture().getDocument().getDocumentGraph().getTokens().get(4)));	
	}
}
