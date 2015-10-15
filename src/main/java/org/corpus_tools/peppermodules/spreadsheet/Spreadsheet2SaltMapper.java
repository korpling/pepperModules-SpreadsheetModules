package org.corpus_tools.peppermodules.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.salt.SaltFactory;
import org.eclipse.emf.common.util.URI;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *  
 * @author Vivian Voigt
 *
 */
public class Spreadsheet2SaltMapper extends PepperMapperImpl implements PepperMapper{
	/**
	 * 
	 */
	private static final Logger logger= LoggerFactory.getLogger(Spreadsheet2SaltMapper.class);
	
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * If you need to make any adaptations to the corpora like adding further meta-annotation, do it here.
	 * When whatever you have done successful, return the status {@link DOCUMENT_STATUS#COMPLETED}. If anything
	 * went wrong return the status {@link DOCUMENT_STATUS#FAILED}.
	 * <br/>
	 * In our dummy implementation, we just add a creation date to each corpus. 
	 */
	@Override
	public DOCUMENT_STATUS mapSCorpus() {
		//getScorpus() returns the current corpus object.
		getCorpus().createMetaAnnotation(null, "date", "1989-12-17");
		return(DOCUMENT_STATUS.COMPLETED);
	}
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * This is the place for the real work. Here you have to do anything necessary, to map a corpus to Salt.
	 * These could be things like: reading a file, mapping the content, closing the file, cleaning up and so on.
	 * <br/>
	 * In our dummy implementation, we do not read a file, for not making the code too complex. We just show how
	 * to create a simple document-structure in Salt, in following steps:
	 * <ol>
	 * 	<li>creating primary data</li>
	 *  <li>creating tokenization</li>
	 *  <li>creating part-of-speech annotation for tokenization</li>
	 *  <li>creating information structure annotation via spans</li>
	 *  <li>creating anaphoric relation via pointing relation</li>
	 *  <li>creating syntactic annotations</li>
	 * </ol>
	 */
	@Override
	public DOCUMENT_STATUS mapSDocument() {
		
		//the method getSDocument() returns the current document for creating the document-structure
		getDocument().setDocumentGraph(SaltFactory.createSDocumentGraph());
		//to get the exact resource, which be processed now, call getResources(), make sure, it was set in createMapper()  
		URI resource= getResourceURI();
		
		//we record, which file currently is imported to the debug stream, in this dummy implementation the resource is null 
		logger.debug("Importing the file {}.", resource);
		
		logger.info(resource.toString());
		
		try {
			File xlsxFile = new File(resource.path());
			InputStream xlsxFileStream = new FileInputStream(xlsxFile);
			XSSFWorkbook xlsxWorkbook = new XSSFWorkbook(xlsxFileStream);
			XSSFSheet sheet = xlsxWorkbook.getSheetAt(0);
			System.out.println(sheet.getRow(0).getCell(0).toString());
			

			xlsxWorkbook.close();
		} catch ( IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		/**
		 * STEP 1:
		 * we create the primary data and hold a reference on the primary data object
		 */
	
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
	
		//now we are done and return the status that everything was successful
		return(DOCUMENT_STATUS.COMPLETED);
	}
	public void setProperties(
			SpreadsheetImporterProperties spreadsheetImporterProperties) {
		// TODO Auto-generated method stub
		
	}	
}