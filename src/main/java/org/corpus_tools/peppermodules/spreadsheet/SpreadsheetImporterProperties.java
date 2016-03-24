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
	public static final String PROP_ANNO_REFERS_TO = "annoPrimRel";
	public static final String PROP_SET_LAYER = "setLayer";

	public SpreadsheetImporterProperties() {
		addProperty(new PepperModuleProperty<>(PROP_PRIMARY_TEXT, String.class, "Defines the name of the column(s), that hold the primary text, this can either be a single column name, or a comma seperated enumeration of column names (key: 'primText', default value: 'tok').", "tok", false));
		addProperty(new PepperModuleProperty<>(PROP_CORPUS_SHEET, String.class, "Defines the sheet, that holds the actual corpus (key: 'primText', default is 'Tabelle1'). If you allways want to use the first sheet of your spreadsheets independent of the name, type in 'FIRST_SHEET' as a value", "Tabelle1", false));
		addProperty(new PepperModuleProperty<>(PROP_CONCATENATE_TEXT, Boolean.class, "Defines, if the textual data shall be concatenated or if a new string object shall be created (key: 'concatenateText', default is 'true').", true, false));
		addProperty(new PepperModuleProperty<>(PROP_ANNO_REFERS_TO, String.class, "Defines which annotation layer refers to which primary text layer, therefor the annotation layer name is followed by the name of the layer, that holds the primary text in square brackets. A possible key-value set could be: key='annoPrimRel', value='anno1[tok1]' (key: 'annoPrimRel', default is 'null').", null, false));
		addProperty(new PepperModuleProperty<>(PROP_SET_LAYER, String.class, "Defines the corresponding layer to the given annotation layer (key: 'setLayer') ", false));
			
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
	
	public String getAnnoPrimRel() {
		return (String) getProperty(PROP_ANNO_REFERS_TO).getValue();
	}
	
	public String getLayer(){
		return (String) getProperty(PROP_SET_LAYER).getValue();
	}
}
