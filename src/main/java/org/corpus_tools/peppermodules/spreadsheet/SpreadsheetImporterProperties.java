package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;


public class SpreadsheetImporterProperties extends PepperModuleProperties {

	/**
	 * 
	 */
	private static final long serialVersionUID = 6193481351811254201L;
	public static final String PROP_CONCATENATE_TEXT = "concatenateText";
	public static final String PROP_PRIMARY_TEXT = "primText";
	public static final String PROP_CORPUS_SHEET = "corpusSheet";

	public SpreadsheetImporterProperties() {
		addProperty(new PepperModuleProperty<>(PROP_PRIMARY_TEXT, String.class, "Defines the name of the xml tag that includes the textual data to be used as primary text (key: 'textElement', default value: 'unicode').", "tok"));
		addProperty(new PepperModuleProperty<>(PROP_CORPUS_SHEET, String.class, "Defines the sheet, that holds the actual corpus (key: 'primText', default is 'Tabelle1'). If you allways want to use the first sheet of your spreadsheets independent of the name, type in 'FIRST_SHEET' as a value", "Tabelle1", false));
		addProperty(new PepperModuleProperty<>(PROP_CONCATENATE_TEXT, Boolean.class, "Defines, if the textual data shall be concatenated or if a new string object shall be created (key: 'concatenateText', default is 'true').", true, false));
		}

	public Boolean concatenateText() {
		return (Boolean) getProperty(PROP_CONCATENATE_TEXT).getValue();
	}
	
	public String getPrimaryText() {
		return (String) getProperty(PROP_PRIMARY_TEXT).getValue();
	}

	public String getCorpusSheet() {
		
		return (String) getProperty(PROP_CORPUS_SHEET).getValue();
	}
}
