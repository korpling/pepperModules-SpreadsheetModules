package org.corpus_tools.peppermodules.spreadsheet.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;

import javax.xml.parsers.ParserConfigurationException;

import org.corpus_tools.pepper.testFramework.PepperTestUtil;
import org.corpus_tools.peppermodules.spreadsheet.Spreadsheet2SaltMapper;
import org.corpus_tools.peppermodules.spreadsheet.SpreadsheetImporterProperties;
import org.corpus_tools.salt.SaltFactory;
import org.eclipse.emf.common.util.URI;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class Spreadsheet2SaltMapperTest {
	
	private Spreadsheet2SaltMapper fixture = null;
	ByteArrayOutputStream outStream = null;
	
	OutputStream outFile = null;
	
	public Spreadsheet2SaltMapper getFixture() {
		return fixture;
	}

	public void setFixture(Spreadsheet2SaltMapper fixture) {
		this.fixture = fixture;
	}

	@Before
	public void setUp() throws FileNotFoundException {
		outStream = new ByteArrayOutputStream();
		String inputFile = PepperTestUtil.getTestResources() + "ContrafaytKreuterbuch_1532.xlsx";
		File testFile = new File(inputFile);
		outFile = new FileOutputStream(testFile);
		setFixture(new Spreadsheet2SaltMapper());
		getFixture().setDocument(SaltFactory.createSDocument());
		getFixture().getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());
		getFixture().getDocument().setName("sample");
		getFixture().setProperties(new SpreadsheetImporterProperties());
	}

	@After
	public void tearDown() {
		outStream.reset();
	}

	private void start(Spreadsheet2SaltMapper mapper, String xlsxString) throws ParserConfigurationException, IOException {
		File tmpDir = new File(System.getProperty("java.io.tmpdir") + "/xlsx2saltTest/");
		tmpDir.mkdirs();
		File tmpFile = new File(tmpDir.getAbsolutePath() + System.currentTimeMillis() + ".xlsx");
		PrintWriter writer = null;
		try {
			writer = new PrintWriter(tmpFile, "UTF-8");
			writer.println(xlsxString);
		} finally {
			if (writer != null)
				writer.close();
		}

		this.getFixture().setResourceURI(URI.createFileURI(tmpFile.getAbsolutePath()));
		this.getFixture().mapSDocument();
	}

	/**
	 * test for correct generation of primary data with concatenated STextualDS
	 * 
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	@SuppressWarnings("unchecked")
	@Test
	public void testPrimDataConcatenate() throws ParserConfigurationException, IOException {

		start(getFixture(), outStream.toString());

		assertEquals(1, getFixture().getDocument().getDocumentGraph().getTextualDSs().size());
		assertNotNull(getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0));
		assertEquals("", getFixture().getDocument().getDocumentGraph().getTextualDSs().get(0).getText());
	}
}
